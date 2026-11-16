"""確認の証拠の計算：ページの部品、表示環境、表示の証拠、内容の証拠、資料全体の証拠、診断の証拠の範囲、画像の比較、画素のハッシュ値。

PowerPoint と確認記録（`records`）に依存しない。PDF の読み取り（`slidekit.pdf`）と Pillow は関数の中で読み込む。
ページ画像の作成（`render_page_image`）は、子プロセスでの呼び出しに対応するため、引数と戻り値を pickle できる値だけにする。
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
import hashlib
import json
import math
from pathlib import Path
import posixpath
import re
from typing import Any

from lxml import etree

from .files import file_hash, json_digest
from .ooxml import EMU_PER_PT, NS, read_package, slide_order, tag

EVIDENCE_VERSION = 2
# 画像化の倍率。生成の画像化と表示環境の証拠で共有する。
RENDER_SCALE = 2
# 互換表示の確認用複製の作り方（fallback.prepare_fallback）を変えたら上げる。
FALLBACK_TRANSFORM_VERSION = 1
# 引き継ぐときの画像の比較（render_difference）。
RENDER_CHECK_METHOD = 1
NOISE_LEVEL = 64
NOISE_PIXELS = 16
MASK_PADDING_PT = 2.0
# 一覧画像に並べる縮小画像の最大の寸法（画素）。
THUMBNAIL_SIZE = (480, 270)

FOOTER_ID = "_theme/footer"
TITLE_ID = "_theme/title"
_SECTION_EXTENSION = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
_RELATIONSHIP_PREFIXES = tuple(
    "{%s}" % namespace for namespace in (NS["r"], "http://purl.oclc.org/ooxml/officeDocument/relationships")
)
# 構造をたどる部品の種類（関係の種類の末尾の名前）。それ以外の参照先は内容の SHA-256 にする。
_STRUCTURED = frozenset({"slide", "slideLayout", "slideMaster", "theme"})
# 属性から参照されてもたどらない関係。ノートとコメントは表示に入らず、マスターからレイアウトへの参照は循環する。
_UNFOLLOWED = {"slide": frozenset({"notesSlide", "comments"}), "slideMaster": frozenset({"slideLayout"})}
# 属性から参照されない関係のうち、証拠に入れるもの（部品の種類ごとに 1 種類）。
_IMPLICIT = {"slide": "slideLayout", "slideLayout": "slideMaster", "slideMaster": "theme"}
_PAGE_VALUES = ("size", "notes", "content_area")
_ELEMENT_MODES = ("all", "content")


@dataclass(frozen=True)
class PageParts:
    """1 ページの表示を決める PPTX の部品の証拠。"""

    slide_id: str
    number: int  # PPTX の中の順（1 始まり）
    choice: str  # PowerPoint の標準表示用。mc:AlternateContent を最初の mc:Choice の中身に置き換える
    full: str  # 互換画像（mc:Fallback）を含む全体
    footer: dict | None  # {"area": [x, y, w, h], "text": "12"}。_theme/footer がなければ None
    fields: tuple[str, ...]  # 表示されうるフィールド（a:fld）の type の一覧
    # ページ番号の図形（_theme/footer）を除いた choice と full。ページ番号の枠・書式・文字だけが変わったページを、
    # 番号の領域を除いた画像の比較で引き継ぐのに使う（footer がなければ None）。
    body_choice: str | None = None
    body_full: str | None = None

    def to_json(self) -> dict:
        """生成記録の page_parts の値（slide_id と number は含めない）。"""
        value = {
            "choice": self.choice,
            "full": self.full,
            "footer": None if self.footer is None else dict(self.footer),
            "fields": list(self.fields),
        }
        if self.body_choice is not None:
            value.update(body_choice=self.body_choice, body_full=self.body_full)
        return value


def _parser():
    # lxml の解析器はスレッドの間で共有できないので、呼び出しごとに作る。
    return etree.XMLParser(resolve_entities=False, no_network=True)


def _relationships(parts: Mapping[str, bytes], name: str, parser) -> dict[str, tuple[str, str | None, str | None]]:
    """部品の関係を識別子から（種類、パッケージ内の参照先、外部の参照先）への辞書で返す。"""
    path = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels")
    if path not in parts:
        return {}
    result = {}
    for node in etree.fromstring(parts[path], parser):
        if not isinstance(node.tag, str):
            continue
        kind = node.get("Type", "").rsplit("/", 1)[-1]
        target = node.get("Target", "")
        if node.get("TargetMode") == "External":
            result[node.get("Id")] = (kind, None, target)
        else:
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(name), target)).lstrip("/")
            result[node.get("Id")] = (kind, resolved, None)
    return result


def _tree_digest(value) -> str:
    """文字列と列だけの構造の SHA-256。files.json_digest と同じ値になる（非有限の数の置き換えが要らないので速い）。"""
    text = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


_ALTERNATE = tag("mc", "AlternateContent")
_CHOICE = tag("mc", "Choice")
_FALLBACK = tag("mc", "Fallback")


def _structures(node, reference) -> tuple[list, list]:
    """接頭辞・属性の順・関係の識別子によらない構造を、full と choice の両方について 1 回の走査で作る。

    要素ごとに [タグ（名前空間の URI 付き）, 並べた属性, 文字, 子の列, 空白でない tail] とする。
    関係を指す属性は reference で参照先の証拠に置き換える。戻り値は（full の構造、choice でこの要素を置き換える構造の列）で、
    choice では mc:AlternateContent を最初の mc:Choice（なければ mc:Fallback）の子に置き換える（p:grpSp の中も含む）。
    choice で変わらない部分は full と同じ列を共有する。
    """
    attributes = node.items()
    attributes.sort()
    attributes = [
        [name, reference(value)] if name.startswith(_RELATIONSHIP_PREFIXES) else [name, value]
        for name, value in attributes
    ]
    full_children, choice_children, branches = [], [], {}
    shared = True
    for child in node:
        if not isinstance(child.tag, str):
            continue
        full, replacement = _structures(child, reference)
        full_children.append(full)
        if len(replacement) == 1 and replacement[0] is full:
            choice_children.append(full)
        else:
            shared = False
            choice_children.extend(replacement)
        if child.tag in (_CHOICE, _FALLBACK):
            branches.setdefault(child.tag, replacement[0])
    text = node.text or ""
    tail = node.tail if node.tail and node.tail.strip() else ""

    def build(children):
        # 子を持つ要素の空白だけの文字は書式の揺れなので入れない（a:t などの葉の文字はそのまま）。
        return [node.tag, attributes, "" if children and not text.strip() else text, children, tail]

    full = build(full_children)
    if node.tag == _ALTERNATE:
        branch = branches.get(_CHOICE, branches.get(_FALLBACK))
        return full, list(branch[3]) if branch is not None else []
    return full, [full] if shared else [build(choice_children)]


def _unwrap_alternate_content(root) -> None:
    """mc:AlternateContent を最初の mc:Choice（なければ mc:Fallback）の子で置き換える。p:grpSp の中も含む。"""
    for alternative in list(root.iter(_ALTERNATE)):
        parent = alternative.getparent()
        if parent is None:
            continue
        chosen = alternative.find("mc:Choice", NS)
        if chosen is None:
            chosen = alternative.find("mc:Fallback", NS)
        position = parent.index(alternative)
        parent.remove(alternative)
        for offset, child in enumerate(list(chosen) if chosen is not None else []):
            parent.insert(position + offset, child)


def _footer_shapes(root):
    for properties in root.iter(tag("p", "cNvPr")):
        if properties.get("name") == FOOTER_ID and properties.getparent() is not None:
            shape = properties.getparent().getparent()
            if shape is not None:
                yield shape


def _drop_footer(root) -> None:
    """ページ番号の図形を取り除く（ページ番号を除いた証拠に使う）。"""
    for shape in list(_footer_shapes(root)):
        parent = shape.getparent()
        if parent is not None:
            parent.remove(shape)


def _mask_page_number(root) -> None:
    """ページ番号の図形の数字をすべて 0 にする。桁数が同じ番号の変化では証拠を変えない。"""
    for shape in _footer_shapes(root):
        for text in shape.iter(tag("a", "t")):
            text.text = re.sub(r"\d", "0", text.text or "")


def _footer(root) -> dict | None:
    for shape in _footer_shapes(root):
        offset, extent = shape.find("p:spPr/a:xfrm/a:off", NS), shape.find("p:spPr/a:xfrm/a:ext", NS)
        area = None
        if offset is not None and extent is not None:
            area = [
                int(offset.get("x", "0")) / EMU_PER_PT,
                int(offset.get("y", "0")) / EMU_PER_PT,
                int(extent.get("cx", "0")) / EMU_PER_PT,
                int(extent.get("cy", "0")) / EMU_PER_PT,
            ]
        return {"area": area, "text": "".join(node.text or "" for node in shape.iter(tag("a", "t")))}
    return None


def _in_placeholder(node) -> bool:
    return any(ancestor.find("./*/p:nvPr/p:ph", NS) is not None for ancestor in node.iterancestors())


def _fields(root, *, displayed_only: bool) -> set[str]:
    """フィールドの type を集める。displayed_only なら配置用の図形（p:ph を持つ図形）の中を除く。"""
    return {
        field.get("type", "") for field in root.iter(tag("a", "fld")) if not (displayed_only and _in_placeholder(field))
    }


class _Package:
    """部品の証拠を、部品と種類ごとに 1 回だけ（full と choice を同時に）計算する。"""

    def __init__(self, parts: Mapping[str, bytes], *, drop_footer: bool = False):
        self.parts = parts
        # 真なら、ページ番号の図形の数字を 0 にする代わりに、図形ごと除く。
        self._footer = _drop_footer if drop_footer else _mask_page_number
        self._parser = _parser()
        self._relations: dict[str, dict] = {}
        self._memo: dict[tuple[str, str], tuple[str, str]] = {}
        self._contents: dict[str, str] = {}
        self._names: dict[str, str] = {}
        self._active: set[tuple[str, str]] = set()

    def relationships(self, name: str) -> dict:
        if name not in self._relations:
            self._relations[name] = _relationships(self.parts, name, self._parser)
        return self._relations[name]

    def root(self, name: str):
        return etree.fromstring(self.parts[name], self._parser)

    def related(self, name: str, kind: str) -> str | None:
        """name が種類 kind の関係で参照する最初の部品。"""
        for relation_kind, target, external in self.relationships(name).values():
            if relation_kind == kind and external is None and target in self.parts:
                return target
        return None

    def content(self, name: str) -> str:
        if name not in self._contents:
            self._contents[name] = hashlib.sha256(self.parts[name]).hexdigest() if name in self.parts else "missing"
        return self._contents[name]

    def slide_name(self, name: str) -> str:
        # スライドどうしの参照（ページへのリンク）は、参照先のページの識別子で表し、循環を避ける。
        if name not in self._names:
            if name not in self.parts:
                self._names[name] = "missing"
            else:
                slide = self.root(name).find("p:cSld", NS)
                self._names[name] = slide.get("name", "") if slide is not None else ""
        return self._names[name]

    def reference(self, source_kind: str, relation) -> tuple[str, str]:
        """関係を指す属性の値を、参照先の証拠に置き換える。戻り値は（full、choice）の値。"""
        if relation is None:
            return "missing", "missing"
        kind, target, external = relation
        if kind in _UNFOLLOWED.get(source_kind, ()):
            value = f"unfollowed:{kind}"
        elif external is not None:
            value = f"external:{external}"
        elif kind == "slide":
            value = "slide:" + self.slide_name(target)
        elif kind in _STRUCTURED:
            full, choice = self.part(target, kind)
            return "part:" + full, "part:" + choice
        else:
            value = "content:" + self.content(target)
        return value, value

    def part(self, name: str, kind: str) -> tuple[str, str]:
        """部品の構造の証拠（full、choice）。たどる関係の参照先を再帰して含める。"""
        key = (name, kind)
        if key in self._memo:
            return self._memo[key]
        if name not in self.parts:
            return "missing", "missing"
        if key in self._active:
            return f"cycle:{kind}", f"cycle:{kind}"
        self._active.add(key)
        try:
            root = self.root(name)
            self._footer(root)
            relations = self.relationships(name)
            referenced: set[str] = set()
            choice_values: dict[str, str] = {}

            def reference(identifier):
                referenced.add(identifier)
                full, choice = self.reference(kind, relations.get(identifier))
                if choice != full:
                    choice_values[identifier] = choice
                return full

            full_tree, replacement = _structures(root, reference)
            choice_tree = replacement[0]
            if choice_values:
                # 属性から参照する部品の choice と full が違うときだけ、choice の値で作り直す（雛形の部品どうしの参照）。
                unwrapped = self.root(name)
                self._footer(unwrapped)
                _unwrap_alternate_content(unwrapped)
                choice_tree = _structures(
                    unwrapped, lambda identifier: self.reference(kind, relations.get(identifier))[1]
                )[0]
            implicit = [
                (relation_kind, self.part(target, relation_kind))
                for identifier, (relation_kind, target, external) in relations.items()
                if identifier not in referenced and external is None and _IMPLICIT.get(kind) == relation_kind
            ]
            value = tuple(
                _tree_digest(
                    {
                        "tree": tree,
                        "implicit": sorted(
                            [relation_kind, "part:" + digests[index]] for relation_kind, digests in implicit
                        ),
                    }
                )
                for index, tree in enumerate((full_tree, choice_tree))
            )
        finally:
            self._active.discard(key)
        self._memo[key] = value
        return value

    def presentation(self) -> str:
        """全ページに共通の ppt/presentation.xml の値（ページ・マスター・ノートの一覧とセクションを除く）。"""
        name = "ppt/presentation.xml"
        root = self.root(name)
        for local in ("sldIdLst", "sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "custShowLst"):
            for node in root.findall(f"p:{local}", NS):
                root.remove(node)
        for extension in root.findall("p:extLst/p:ext", NS):
            if extension.get("uri") == _SECTION_EXTENSION:
                extension.getparent().remove(extension)
        relations = self.relationships(name)
        tree = _structures(root, lambda identifier: self.reference("presentation", relations.get(identifier))[0])[0]
        return _tree_digest(tree)


def pptx_pages(path: str | Path) -> dict[str, PageParts]:
    """PPTX の各ページの部品の証拠を、ページの識別子（p:cSld/@name）から PageParts への辞書で返す（表示順）。

    部品名・画像名・関係の識別子・文書情報・ノートによらない。参照する画像などは内容の SHA-256 で入れる。
    choice は互換画像（mc:Fallback だけが参照するメディア）を含まないので、互換画像の中身だけが違う PPTX では一致する。
    """
    parts = read_package(path)
    ordered, _ = slide_order(parts)
    package = _Package(parts)
    bodies = _Package(parts, drop_footer=True)
    presentation = package.presentation()
    result: dict[str, PageParts] = {}
    for number, name in enumerate(ordered, 1):
        root = package.root(name)
        slide = root.find("p:cSld", NS)
        slide_id = slide.get("name") if slide is not None else None
        if not slide_id or slide_id in result:
            raise ValueError(f"PPTX のページの識別子が空か重複している: {number} ページ目")
        fields = _fields(root, displayed_only=False)
        layout = package.related(name, "slideLayout")
        master = package.related(layout, "slideMaster") if layout else None
        for shared in (layout, master):
            if shared:
                fields |= _fields(package.root(shared), displayed_only=True)
        full, choice = package.part(name, "slide")
        footer = _footer(root)
        body = bodies.part(name, "slide") if footer is not None else None
        result[slide_id] = PageParts(
            slide_id=slide_id,
            number=number,
            choice=json_digest({"slide": choice, "presentation": presentation}),
            full=json_digest({"slide": full, "presentation": presentation}),
            footer=footer,
            fields=tuple(sorted(fields)),
            body_choice=None if body is None else json_digest({"slide": body[1], "presentation": presentation}),
            body_full=None if body is None else json_digest({"slide": body[0], "presentation": presentation}),
        )
    return result


def display_environment(
    renderer: str,
    detected: Mapping | None,
    actual: Mapping | None,
    fonts: Mapping | None,
    rasterizer: str,
    *,
    scale: float = RENDER_SCALE,
) -> str:
    """表示環境の証拠。

    detected は事前確認の結果（版、実行ファイル、Windows の利用者、フォントの照合の結果）、actual は変換の結果
    （表示ソフト、版、ビルド、実行ファイル）、fonts は実行環境で宣言したフォントごとの内容の証拠である。
    状態のフォルダー・登録された書体の一覧・変換ごとに変わる値（所要時間・入出力のハッシュ値）と、
    計測で使ったフォントの一覧は入れない（使うフォントが増えただけで全ページの証拠が変わるのを避ける）。
    """
    detected, actual = detected or {}, actual or {}
    keys = ("family", "variant", "table_sha256", "windows_path", "windows_index")
    matches = sorted(
        ({key: match.get(key) for key in keys} for match in detected.get("font_matches") or []), key=json_digest
    )
    return json_digest(
        {
            "format": EVIDENCE_VERSION,
            "renderer": renderer,
            "detected": {
                "font_matches": matches,
                **{key: detected.get(key) for key in ("version", "executable", "windows_user_sid")},
            },
            "actual": {key: actual.get(key) for key in ("renderer", "version", "build", "executable")},
            "fonts": dict(fonts or {}),
            "rasterizer": rasterizer,
            "scale": float(scale),
        }
    )


def visual_body_digest(
    page: PageParts, renderer: str, variant: str, environment_hash: str, *, render_date: str, number: int | None = None
) -> str | None:
    """ページ番号の図形を除いた表示の証拠（visual_digest と同じ形）。ページ番号の図形がなければ None。

    確認の引き継ぎで、ページ番号の枠・書式・文字だけが変わったページを見分けるのに使う。確認対象の証拠には使わない。
    """
    if page.body_choice is None or page.body_full is None:
        return None
    body = PageParts(page.slide_id, page.number, page.body_choice, page.body_full, None, page.fields)
    return visual_digest(body, renderer, variant, environment_hash, render_date=render_date, number=number)


def visual_digest(
    page: PageParts, renderer: str, variant: str, environment_hash: str, *, render_date: str, number: int | None = None
) -> str:
    """表示の証拠。PowerPoint の標準表示は choice、それ以外（LibreOffice、互換表示）は full を使う。

    number は変換した資料の中のページの位置である。標準表示では省略でき、page.number を使う。
    互換表示の確認用複製は数式のページだけに間引くので、複製の中の位置（prepare_fallback の kept_slides）を必ず渡す。
    フィールドがあるページだけ位置を入れ、日付のフィールドがあれば render_date（生成した日。UTC の YYYY-MM-DD）も入れる。
    """
    if variant not in ("standard", "fallback"):
        raise ValueError(f"表示方法は standard か fallback にする: {variant}")
    if number is None:
        if variant == "fallback":
            raise ValueError("互換表示の証拠には、確認用複製の中のページの位置（kept_slides の number）を渡す")
        number = page.number
    parts = page.choice if (renderer == "powerpoint" and variant == "standard") else page.full
    value: dict[str, Any] = {
        "format": EVIDENCE_VERSION,
        "parts": parts,
        "renderer": renderer,
        "variant": variant,
        "environment": environment_hash,
    }
    if variant == "fallback":
        value["transform"] = FALLBACK_TRANSFORM_VERSION
    if page.fields:
        # 番号のフィールドがあると、資料の中の位置で表示が変わる。
        value["number"] = number
    if any(kind.startswith("datetime") for kind in page.fields):
        value["date"] = render_date
    return json_digest(value)


def overview_digest(contact_hashes: Sequence[str], environment_hash: str, deck_hash: str) -> str:
    """一覧画像の確認の証拠：一覧画像の画素、表示環境、資料全体の証拠（deck_digest）。"""
    return json_digest(
        {
            "format": EVIDENCE_VERSION,
            "contacts": list(contact_hashes),
            "environment": environment_hash,
            "deck": deck_hash,
        }
    )


def pixel_digest(width: int, height: int, samples: bytes) -> str:
    """画像の画素のハッシュ値。samples は RGB の画素を行の順に並べたバイト列（Pillow の tobytes）。"""
    if len(samples) != width * height * 3:
        raise ValueError(f"画素の数が寸法と合わない: {width}x{height}、{len(samples)} バイト")
    digest = hashlib.sha256(f"{width}x{height}:RGB:".encode())
    digest.update(samples)
    return digest.hexdigest()


def render_page_image(pdf: str, index: int, path: str, scale: float = RENDER_SCALE, *, thumbnail: bool = True) -> dict:
    """PDF の 1 ページ（index は 0 始まり）を PNG にし、画素のハッシュ値を返す。子プロセスで呼ぶことができる。

    thumbnail が真なら、一覧画像に並べる縮小画像（RGB の画素のバイト列と寸法）も返す。互換表示では一覧画像を作らないので偽にする。
    戻り値：{"path", "pixel_hash"}、thumbnail が真なら {"thumbnail": bytes, "thumbnail_size": [w, h]} を加える。
    """
    from .pdf import open_pdf

    with open_pdf(pdf) as document:
        if not 0 <= index < len(document):
            raise ValueError(f"PDF にそのページがない: {pdf}（{index + 1} ページ目、全 {len(document)} ページ）")
        image = document[index].render(scale)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    image.save(path, "PNG")
    samples = image.tobytes()
    result: dict[str, Any] = {"path": path, "pixel_hash": pixel_digest(image.width, image.height, samples)}
    if thumbnail:
        image.thumbnail(THUMBNAIL_SIZE)
        result.update(thumbnail=image.tobytes(), thumbnail_size=list(image.size))
    return result


def _channel_max(image):
    from PIL import ImageChops

    red, green, blue = image.split()
    return ImageChops.lighter(ImageChops.lighter(red, green), blue)


def _strong(image):
    return image.point(lambda value: 255 if value > NOISE_LEVEL else 0)


def render_difference(
    source: str | Path, target: str | Path, *, masks: Sequence[Sequence[float]] = (), scale: float = RENDER_SCALE
) -> dict:
    """2 つのページ画像を、1 画素のずれを許して比べる。引き継ぐときの安全網に使う。

    戻り値：{"method", "strong_pixels", "size_mismatch"}。寸法が違えば strong_pixels は None。
    masks（pt の [x, y, w, h]）は MASK_PADDING_PT だけ広げ、両方の画像で同じ色に塗ってから比べる。
    数字 1 字程度の小さな変化は見つけることができないので、同一性は部品の証拠で確かめる。
    呼び出し側は、画素のハッシュ値が同じなら比較を省いて strong_pixels を 0 とする。
    """
    from PIL import Image, ImageChops, ImageDraw, ImageFilter

    with Image.open(source) as first, Image.open(target) as second:
        before, after = first.convert("RGB"), second.convert("RGB")
    result = {"method": RENDER_CHECK_METHOD, "strong_pixels": None, "size_mismatch": before.size != after.size}
    if result["size_mismatch"]:
        return result
    for mask in masks:
        valid = len(mask) == 4 and all(isinstance(value, (int, float)) and math.isfinite(value) for value in mask)
        if not valid or min(mask[2:]) < 0:
            raise ValueError(f"比較から除く領域は有限の [x, y, w, h] にする: {mask}")
        x, y, width, height = mask
        box = [
            math.floor((x - MASK_PADDING_PT) * scale),
            math.floor((y - MASK_PADDING_PT) * scale),
            math.ceil((x + width + MASK_PADDING_PT) * scale),
            math.ceil((y + height + MASK_PADDING_PT) * scale),
        ]
        for image in (before, after):
            ImageDraw.Draw(image).rectangle(box, fill=(255, 0, 255))
    box = _strong(_channel_max(ImageChops.difference(before, after))).getbbox()
    if box is None:
        result["strong_pixels"] = 0
        return result
    left, top, right, bottom = box
    box = (max(left - 2, 0), max(top - 2, 0), min(right + 2, before.width), min(bottom + 2, before.height))
    before, after = before.crop(box), after.crop(box)
    worst = 0
    for image, other in ((before, after), (after, before)):
        # 相手の画像で同じ位置の周囲 3×3 画素の最小値〜最大値の範囲から、NOISE_LEVEL を超えて外れる画素の数を求める。
        low, high = other.filter(ImageFilter.MinFilter(3)), other.filter(ImageFilter.MaxFilter(3))
        excess = ImageChops.lighter(ImageChops.subtract(low, image), ImageChops.subtract(image, high))
        worst = max(worst, _strong(_channel_max(excess)).histogram()[255])
    result["strong_pixels"] = worst
    return result


def difference_boxes(
    source: str | Path,
    target: str | Path,
    *,
    masks: Sequence[Sequence[float]] = (),
    scale: float = RENDER_SCALE,
    block_pt: float = 12.0,
) -> list[list[float]]:
    """2 つのページ画像で、NOISE_LEVEL を超えて色が違う領域（pt の [x, y, w, h]）の列。比較画像に枠を描くのに使う。

    画像を block_pt 四方の区画に分け、強い差のある区画を、隣り合うものどうしまとめた外接矩形にする。masks の領域
    （ページ番号の枠など）は render_difference と同じく除く。寸法が違う画像はページ全体を 1 つの領域にする。
    引き継ぎの判定には使わない（判定は render_difference の 1 画素のずれを許した比較で行う）。
    """
    from PIL import Image, ImageChops, ImageDraw

    with Image.open(source) as first, Image.open(target) as second:
        before, after = first.convert("RGB"), second.convert("RGB")
    if before.size != after.size:
        return [[0.0, 0.0, after.width / scale, after.height / scale]]
    for mask in masks:
        x, y, width, height = (float(value) for value in mask)
        box = [
            math.floor((x - MASK_PADDING_PT) * scale),
            math.floor((y - MASK_PADDING_PT) * scale),
            math.ceil((x + width + MASK_PADDING_PT) * scale),
            math.ceil((y + height + MASK_PADDING_PT) * scale),
        ]
        for image in (before, after):
            ImageDraw.Draw(image).rectangle(box, fill=(255, 0, 255))
    strong = _strong(_channel_max(ImageChops.difference(before, after)))
    if strong.getbbox() is None:
        return []
    size = max(1, round(block_pt * scale))
    columns, rows = math.ceil(strong.width / size), math.ceil(strong.height / size)
    marked = {
        (column, row)
        for row in range(rows)
        for column in range(columns)
        if strong.crop((column * size, row * size, (column + 1) * size, (row + 1) * size)).getbbox() is not None
    }
    boxes = []
    while marked:
        stack = [marked.pop()]
        group = set(stack)
        while stack:
            column, row = stack.pop()
            for neighbour in ((column + dx, row + dy) for dx in (-1, 0, 1) for dy in (-1, 0, 1)):
                if neighbour in marked:
                    marked.remove(neighbour)
                    group.add(neighbour)
                    stack.append(neighbour)
        left = min(column for column, _ in group) * size
        top = min(row for _, row in group) * size
        right = min(strong.width, (max(column for column, _ in group) + 1) * size)
        bottom = min(strong.height, (max(row for _, row in group) + 1) * size)
        boxes.append([left / scale, top / scale, (right - left) / scale, (bottom - top) / scale])
    return sorted(boxes, key=lambda box: (box[1], box[0]))


def _fragments(value) -> list:
    """断片の列（種類・文字列・書式・リンク）。文字列は 1 つの文字の断片とする。"""
    if value is None:
        return []
    if isinstance(value, str):
        return [{"kind": "text", "text": value}]
    return [{"kind": "text", "text": fragment} if isinstance(fragment, str) else dict(fragment) for fragment in value]


def _connection(reference) -> dict | None:
    # 間隔（gap）と辺の上の位置（at）は表示の値なので入れない。
    if not isinstance(reference, Mapping):
        return None
    return {"element_id": reference.get("element_id"), "side": reference.get("side")}


def _image_digest(element: Mapping, project) -> str:
    if element.get("source_sha256"):
        return element["source_sha256"]
    path = element.get("source_path") or element.get("path")
    if path and str(path) in project.input_files:
        return project.input_files[str(path)]
    if path and Path(path).is_file():
        return file_hash(path)
    raise ValueError(f"画像の内容を確かめられない: {element.get('id')}")


def _content_value(slide: Mapping, project) -> dict:
    """ページの内容の証拠：題名・文字・数式・ノート・出典・図の構造と、そのページが参照する元資料・文献項目。

    slide は素材の準備（`assets.prepare_assets`）の後の場面のページである。画像は元素材の内容の SHA-256（`source_sha256`）で
    入れ、実行ごとに変わる変換先のパスは使わない。文献項目の証拠は `project.config["_bibliography_entries"]`（構築が入れる）を使う。
    番号・座標・寸法・要素の書式・説明時間・作成順・テーマの要素（_theme/）・許容と役割の宣言・まとまりの書き出し・
    画像の倍率と埋め込み方・記述規則・文書情報・依頼内容（brief.md）・参照しない元資料と文献は入れない。
    """
    documents = {document.get("id"): document for document in project.config.get("inputs", {}).get("documents", [])}
    entries = project.config.get("_bibliography_entries", {})
    items, used_documents, used_entries = [], {}, {}
    title_content = None
    for element in slide.get("elements", []):
        identifier = str(element.get("id", ""))
        if identifier == TITLE_ID:
            title_content = _fragments(element.get("content"))
        if identifier.startswith("_theme/"):
            continue
        item = {
            "id": identifier,
            "kind": element.get("kind"),
            "content": _fragments(element.get("content")),
            "sources": list(element.get("sources") or []),
        }
        for key in ("role", "semantic", "description", "shape", "parent_id"):
            if element.get(key) is not None:
                item[key] = element[key]
        if element.get("kind") == "image":
            item["image"] = [_image_digest(element, project), element.get("page")]
        if element.get("kind") in ("line", "arrow"):
            item["connections"] = [_connection(element.get("start_ref")), _connection(element.get("end_ref"))]
        tail = element.get("tail")
        target = tail.get("target") if isinstance(tail, Mapping) else None
        if isinstance(target, Mapping) and target.get("element_id") is not None:
            # しっぽで指した要素（tail_toward）。しっぽの辺・位置・寸法は座標と同じく内容に入れない。
            item["tail_target"] = target["element_id"]
        if "row_keys" in element or "column_keys" in element:
            item["table"] = [element.get("row_keys"), element.get("column_keys")]
        for source in item["sources"]:
            if isinstance(source, str):
                source = {"document_id": source} if source in documents else {"bib_key": source}
            if source.get("document_id"):
                document = documents.get(source["document_id"])
                if document and document.get("affects_content") is False:
                    used_documents[source["document_id"]] = "untracked"
                else:
                    used_documents[source["document_id"]] = (
                        project.input_files.get(str(document.get("path")), "missing") if document else "missing"
                    )
            if source.get("bib_key"):
                used_entries[source["bib_key"]] = entries.get(source["bib_key"], "missing")
        items.append(item)
    return {
        "format": EVIDENCE_VERSION,
        "slide_id": slide["id"],
        "title": slide.get("title"),
        "title_content": title_content,
        "notes": slide.get("notes"),
        "elements": sorted(items, key=lambda item: item["id"]),
        "documents": used_documents,
        "bibliography": used_entries,
    }


def content_digest(slide: Mapping, project) -> str:
    """ページの内容の証拠（`_content_value` の値のハッシュ値）。"""
    return json_digest(_content_value(slide, project))


def content_parts(slide: Mapping, project) -> dict:
    """内容の証拠の内訳。page はページ自身（題名・文字・ノート・出典など）のハッシュ値、documents と bibliography は
    参照する元資料と文献項目ごとのハッシュ値である。内容の確認が引き継がれなかった理由（本文の変更か元資料の変更か）を示すのに使う。
    """
    value = _content_value(slide, project)
    page = {key: item for key, item in value.items() if key not in ("documents", "bibliography")}
    return {"page": json_digest(page), "documents": value["documents"], "bibliography": value["bibliography"]}


def deck_digest(scene: Mapping, project) -> str:
    """資料全体の証拠：文書情報・依頼内容（brief.md）・ページの順と題名・ページ寸法。一覧画像の確認の証拠に加える。

    依頼内容の場所は project.brief_path（属性がないか None なら依頼内容なし）。
    """
    brief_path = getattr(project, "brief_path", None)
    brief = None
    if brief_path:
        brief = project.input_files.get(str(brief_path))
        if brief is None and Path(brief_path).is_file():
            brief = file_hash(brief_path)
    return json_digest(
        {
            "format": EVIDENCE_VERSION,
            "metadata": project.config.get("metadata"),
            "brief": brief,
            "slides": [[slide["id"], slide.get("title")] for slide in scene["slides"]],
            "size": scene.get("size"),
        }
    )


def scoped_evidence(
    scope, *, rules: Mapping, roles: Mapping, elements: Sequence[Mapping], slide: Mapping | None, size: Any
) -> dict:
    """診断の種類の証拠の範囲（diagnostics.EvidenceScope）にしたがって、診断の証拠を作る。

    rules は記述規則の全体、roles はテーマの役割の定義、elements は診断の対象の要素（0〜2 個）、slide はページである。
    scope.rules が None なら記述規則の全体、項目の組ならその項目だけを入れる。scope.roles が真なら要素の役割の定義を入れる。
    scope.element が "all" なら要素の全体、"content" なら kind・content・role・semantic だけを入れる。
    scope.page の "size"・"content_area" はページの値を入れる。"notes" はページ単位の診断（要素なし）だけに入れる
    （要素の判定はノートを使わない。 0.1 の証拠と同じ範囲）。フォントの一覧は呼び出し側が scope.fonts にしたがって加える。
    """
    if scope.element not in _ELEMENT_MODES:
        raise ValueError(f"証拠の範囲の element が不正である: {scope.element}")
    unknown = [name for name in scope.page if name not in _PAGE_VALUES]
    if unknown:
        raise ValueError(f"証拠の範囲の page が不正である: {', '.join(unknown)}")
    elements = sorted(elements, key=lambda element: str(element.get("id", "")))
    if scope.element == "content":
        elements = [
            {
                "kind": element.get("kind"),
                "content": _fragments(element.get("content")),
                "role": element.get("role"),
                "semantic": element.get("semantic"),
            }
            for element in elements
        ]
    selected = dict(rules) if scope.rules is None else {key: rules.get(key) for key in scope.rules}
    result: dict[str, Any] = {"rules": selected, "elements": list(elements)}
    if scope.roles:
        used = sorted({element.get("role") for element in elements if element.get("role")})
        result["roles"] = {role: roles.get(role) for role in used}
    page = slide or {}
    if "size" in scope.page:
        result["size"] = size
    if "notes" in scope.page and not elements:
        result["notes"] = page.get("notes")
    if "content_area" in scope.page:
        result["content_area"] = page.get("content_area")
    return result
