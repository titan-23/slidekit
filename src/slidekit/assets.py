"""ローカルの素材を、元ファイルを変えずに配置用画像へ変換する。

PDF・SVG の素材は PNG にする。画素数は配置寸法（pt）× raster_scale（既定 2）で、画像の指定がなければ deck.toml の
[images] の値を使う。SVG の素材は、vector が真なら SVG でも保存し、書き出しが PNG に加えて SVG を埋め込む。
PDF の素材は PNG だけにする（vector を指定した PDF の素材は誤りにする。資料の既定の vector は SVG の素材だけに適用する）。
保存する SVG は、外部の参照と実行される内容のない自己完結したものだけにする。
"""

from __future__ import annotations

import hashlib
import io
import json
import math
from pathlib import Path
import re

from .errors import InputError
from .files import file_hash
from .layout import Rect

# 画像の指定がなく、資料に [images] の設定もない（試験用の簡易な資料など）ときの値。0.1 と同じ。
DEFAULT_RASTER_SCALE = 2
RASTER_SCALE_RANGE = (1, 16)
# PDF の素材に vector を指定したときの説明。PDF を劣化させずに置くには、SVG に変換して使う。
PDF_VECTOR_MESSAGE = (
    "vector は SVG の素材だけに指定できます。PDF の素材は PNG にします。拡大しても劣化させたくない図は、"
    "SVG に変換して使ってください（例：pdftocairo -svg -f 1 -l 1 図.pdf 図.svg）"
)


def _image_options(element: dict, project, suffix: str, slide_id) -> tuple[float, bool]:
    """画像の raster_scale と vector を、画像の指定、deck.toml の [images]、既定値の順に決めて確かめる。"""
    context = {"code": "invalid_image_option", "slide_id": slide_id, "element_id": element.get("id")}
    specified = {name: element.get(name) for name in ("raster_scale", "vector") if element.get(name) is not None}
    if specified and suffix not in {".pdf", ".svg"}:
        raise InputError(f"{'・'.join(specified)} は PDF・SVG の素材だけに指定できます: {element.get('id')}", **context)
    settings = (getattr(project, "config", None) or {}).get("images") or {}
    scale = specified.get("raster_scale", settings.get("raster_scale", DEFAULT_RASTER_SCALE))
    vector = specified.get("vector", settings.get("vector", False))
    low, high = RASTER_SCALE_RANGE
    if (
        isinstance(scale, bool)
        or not isinstance(scale, (int, float))
        or not math.isfinite(scale)
        or not low <= scale <= high
    ):
        raise InputError(f"raster_scale は {low} 以上 {high} 以下の数で指定してください: {scale!r}", **context)
    if not isinstance(vector, bool):
        raise InputError(f"vector は真偽値で指定してください: {vector!r}", **context)
    if suffix == ".pdf":
        if specified.get("vector"):
            raise InputError(f"{PDF_VECTOR_MESSAGE}: {element.get('id')}", **context)
        vector = False
    return scale, vector


def _input_hash(project, source: Path) -> str:
    """resolve_input が記録した素材のハッシュ値（素材を再びハッシュしない）。記録がなければ求める。"""
    recorded = (getattr(project, "input_files", None) or {}).get(str(source))
    return recorded if recorded else file_hash(source)


def _render_svg(data: bytes, pixels: list[int]) -> bytes:
    """自己完結を確かめた SVG を、幅 pixels[0] で PNG にする。高さが pixels[1] に足りなければ高さで合わせ直す。

    変換器の取得関数も data: 以外を拒否し、CSS のエスケープや入れ子の SVG を含めて外部素材の欠落を黙認しない。
    """
    from PIL import Image

    try:
        import cairosvg
    except ImportError as exc:
        raise ValueError("SVG の変換には slidekit[svg] が必要である") from exc

    def embedded_only(url, resource_type):
        if not url.startswith("data:"):
            raise ValueError("SVG の参照素材は事前に埋め込む必要がある")
        return cairosvg.url.fetch(url, resource_type)

    rendered = cairosvg.surface.PNGSurface.convert(bytestring=data, output_width=pixels[0], url_fetcher=embedded_only)
    with Image.open(io.BytesIO(rendered)) as image:
        needs_height = image.height < pixels[1]
    if needs_height:
        rendered = cairosvg.surface.PNGSurface.convert(
            bytestring=data, output_height=pixels[1], url_fetcher=embedded_only
        )
    return rendered


def prepare_assets(scene: dict, project, directory: Path) -> list[dict]:
    """scene の画像参照だけを変換先へ更新し、素材と変換条件を返す。

    記録の scale は使った raster_scale、vector_output・vector_sha256 は SVG でも保存したときのファイルとハッシュ値
    （そうでなければ None）である。SVG でも保存した画像には、場面に vector_path を書く。
    """
    from PIL import Image

    records = []
    for slide in scene["slides"]:
        for element in slide["elements"]:
            if element["kind"] != "image":
                continue
            source = project.resolve_input(element["path"])
            suffix = source.suffix.lower()
            scale, vector = _image_options(element, project, suffix, slide.get("id"))
            area = element["area"]
            pixels = [max(1, round(area[2] * scale)), max(1, round(area[3] * scale))]
            record = {
                "source": str(source),
                "sha256": _input_hash(project, source),
                "page": element.get("page"),
                "target_pixels": pixels,
            }
            # 変換先の名前は 0.1 と同じ項目（素材・ページ・画素数）から決める。
            key = hashlib.sha256(json.dumps(record, sort_keys=True).encode()).hexdigest()
            record.update(scale=scale, vector_output=None, vector_sha256=None)
            svg = None
            if suffix in {".png", ".jpg", ".jpeg"}:
                with Image.open(source) as image:
                    image.verify()
                destination = source
            elif suffix in {".pdf", ".svg"}:
                directory.mkdir(parents=True, exist_ok=True)
                destination = directory / f"{key}.png"
                if suffix == ".pdf":
                    from .pdf import render_to_pixels

                    page_number = element.get("page")
                    if not isinstance(page_number, int) or isinstance(page_number, bool) or page_number < 1:
                        raise ValueError("PDF 素材は page に 1 始まりのページを指定する")
                    render_to_pixels(source, page_number, pixels).save(destination, "PNG")
                else:
                    data = source.read_bytes()
                    _check_svg(data)
                    with Image.open(io.BytesIO(_render_svg(data, pixels))) as image:
                        image.save(destination)
                    if vector:
                        svg = data
            else:
                raise ValueError(f"未対応の素材形式: {source.suffix}")
            if svg is not None:
                vector_path = directory / f"{key}.svg"
                vector_path.write_bytes(svg)
                record.update(vector_output=str(vector_path), vector_sha256=hashlib.sha256(svg).hexdigest())
                element["vector_path"] = str(vector_path)
            with Image.open(destination) as image:
                record["image_size"] = list(image.size)
            element["source_path"] = str(source)
            element["path"] = str(destination)
            element["source_sha256"] = record["sha256"]
            element["image_size"] = record["image_size"]
            frame = Rect(*area)
            element["actual_area"] = (
                list(frame.contain(*record["image_size"])) if element.get("fit", "contain") == "contain" else list(area)
            )
            record["output"] = str(destination)
            records.append(record)
    return records


def _check_svg(data: bytes) -> None:
    """変換器が参照先を取得しないよう、自己完結した SVG のみ受け付ける。"""
    from lxml import etree
    import tinycss2

    def check_css(tokens):
        for token in tokens:
            if token.type == "at-keyword" and token.lower_value == "import":
                raise ValueError("SVG の外部スタイルシートには対応していない")
            target = None
            if token.type == "url":
                target = token.value
            if token.type == "function" and token.lower_name == "url":
                target = "".join(
                    getattr(part, "value", "") for part in token.arguments if part.type != "comment"
                ).strip()
            if target is not None and not target.startswith(("#", "data:")):
                raise ValueError("SVG の参照素材は事前に埋め込む必要がある")
            check_css(getattr(token, "arguments", getattr(token, "content", [])))

    if re.search(rb"<!DOCTYPE|<!ENTITY", data, re.I):
        raise ValueError("SVG の文書型宣言と外部実体には対応していない")
    # huge_tree：埋め込み画像を 1 つの data: 属性に持つ SVG は、libxml2 の既定の上限（1 つの属性・文字列で
    # 10,000,000 文字）を超えうる。実体は展開せず（resolve_entities=False）、libxml2 の実体の増幅の上限も残る。
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
    try:
        root = etree.fromstring(data, parser)
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"SVG を XML として解析できない: {exc}") from exc
    # UTF-16 などの文字コードでは上のバイト列の検索に掛からないので、解析した文書でも確かめる。
    docinfo = root.getroottree().docinfo
    if docinfo.doctype or docinfo.internalDTD is not None:
        raise ValueError("SVG の文書型宣言と外部実体には対応していない")
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        local = etree.QName(node).localname
        if local in {"script", "foreignObject"}:
            raise ValueError(f"SVG の {local} には対応していない")
        for key, value in node.attrib.items():
            if etree.QName(key).localname in {"href", "src"} and not value.startswith(("#", "data:")):
                raise ValueError("SVG の参照素材は事前に埋め込む必要がある")
        for value in [node.text or "", *node.attrib.values()]:
            check_css(tinycss2.parse_component_value_list(value))
            if re.search(r"@import", value, re.I):
                raise ValueError("SVG の外部スタイルシートには対応していない")
            for target in re.findall(r"url\(\s*['\"]?([^)'\"]+)", value, re.I):
                if not target.strip().startswith(("#", "data:")):
                    raise ValueError("SVG の外部参照は取得しない")
