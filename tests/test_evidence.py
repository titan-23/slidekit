"""確認の証拠（ページの部品、表示環境、表示の証拠、画像の比較、内容の証拠、診断の証拠の範囲、ページ画像）を確かめる。"""

from copy import deepcopy
from dataclasses import replace
import hashlib
from pathlib import Path
import pickle
from types import SimpleNamespace

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
import pytest

from pdf_tools import Document
from slidekit.evidence import (
    NOISE_PIXELS,
    PageParts,
    content_digest,
    deck_digest,
    display_environment,
    overview_digest,
    pixel_digest,
    pptx_pages,
    render_difference,
    render_page_image,
    scoped_evidence,
    visual_digest,
)
from slidekit.fallback import compare_package_structure, finalize_math, prepare_final_form
from slidekit.ooxml import NS, read_package, slide_order, tag, write_package, xml_bytes
from slidekit.pdf import open_pdf
from slidekit.writer import write_pptx

DATE = "2026-10-12"
ENVIRONMENT = "e" * 64


def project(tmp_path, **changes):
    config = {"metadata": {"title": "証拠の検査", "author": "", "last_modified_by": ""}, "math": {"functions": []}}
    config.update(changes)
    return SimpleNamespace(config=config, root=tmp_path, resolve_input=lambda value: (tmp_path / value).resolve())


def element(identifier="body", kind="text", content=None, area=(30, 40, 200, 70), **attributes):
    return {
        "id": identifier,
        "kind": kind,
        "area": list(area),
        "style": {
            "size_pt": 24,
            "min_size_pt": 18,
            "font_japanese": "Yu Gothic UI",
            "font_latin": "Arial",
            "font_math": "Cambria Math",
            "color": "17212D",
            "padding_pt": 0,
            "line_spacing": 1.15,
        },
        "content": [{"kind": "text", "text": "本文の文章"}] if content is None else content,
        **attributes,
    }


def footer(number):
    """作成 API が作るページ番号の図形と同じ識別子と形。"""
    return element("_theme/footer", content=[{"kind": "text", "text": str(number)}], area=(660, 510, 36, 22))


def page(identifier, elements, notes="説明します。"):
    return {"id": identifier, "kind": "content", "title": identifier, "notes": notes, "elements": elements}


def scene(pages):
    return {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {"title": "証拠の検査", "author": "", "last_modified_by": ""},
        "slides": pages,
    }


def picture(path, color):
    Image.new("RGB", (40, 30), color).save(path)
    return str(path)


def write(tmp_path, pages, name, **changes):
    path = tmp_path / f"{name}.pptx"
    write_pptx(scene(pages), project(tmp_path, **changes), path)
    return path


def base_pages(tmp_path):
    image = picture(tmp_path / "figure.png", "red")
    return [
        page("a", [element(), footer(1)]),
        page("b", [element("figure", "image", path=image, area=(30, 120, 80, 60)), footer(2)]),
    ]


def slide_media(path, slide_id):
    """ページが参照する画像の部品名と、スライドの部品名。"""
    parts = read_package(path)
    for name in slide_order(parts)[0]:
        if etree.fromstring(parts[name]).find("p:cSld", NS).get("name") == slide_id:
            relations = etree.fromstring(parts[name.replace("slides/", "slides/_rels/") + ".rels"])
            return name, sorted(node.get("Target") for node in relations if node.get("Type").endswith("/image"))
    raise AssertionError(slide_id)


def test_same_scene_gives_same_page_parts(tmp_path):
    pages = base_pages(tmp_path)
    first, second = write(tmp_path, pages, "first"), write(tmp_path, deepcopy(pages), "second")
    # 文書情報の時刻を変える
    parts = read_package(second)
    core = etree.fromstring(parts["docProps/core.xml"])
    for node in core:
        if etree.QName(node).localname in {"created", "modified"}:
            node.text = "2001-02-03T04:05:06Z"
    parts["docProps/core.xml"] = xml_bytes(core)
    write_package(parts, second)
    assert read_package(first)["docProps/core.xml"] != read_package(second)["docProps/core.xml"]
    before, after = pptx_pages(first), pptx_pages(second)
    assert before == after
    assert list(before) == ["a", "b"]
    assert [item.number for item in before.values()] == [1, 2]
    assert before["a"].to_json() == {
        "choice": before["a"].choice,
        "full": before["a"].full,
        "footer": {"area": [660.0, 510.0, 36.0, 22.0], "text": "1"},
        "fields": [],
    }


def test_inserted_page_keeps_parts_of_other_pages(tmp_path):
    pages = base_pages(tmp_path)
    first = write(tmp_path, pages, "first")
    inserted = [
        page(
            "new",
            [element("lead", "image", path=picture(tmp_path / "lead.png", "blue"), area=(300, 300, 40, 30)), footer(1)],
        ),
        deepcopy(pages[0]),
        deepcopy(pages[1]),
    ]
    inserted[1]["elements"][1], inserted[2]["elements"][1] = footer(2), footer(3)
    second = write(tmp_path, inserted, "inserted")
    # 部品名と画像名が変わる
    assert slide_media(first, "b") != slide_media(second, "b")
    before, after = pptx_pages(first), pptx_pages(second)
    for slide_id in ("a", "b"):
        assert after[slide_id].choice == before[slide_id].choice
        assert after[slide_id].full == before[slide_id].full
    assert (after["a"].number, after["a"].footer["text"]) == (2, "2")
    assert (after["b"].number, after["b"].footer["text"]) == (3, "3")


def test_page_number_digit_count_changes_parts(tmp_path):
    nine = pptx_pages(write(tmp_path, [page("a", [element(), footer(9)])], "nine"))["a"]
    ten = pptx_pages(write(tmp_path, [page("a", [element(), footer(10)])], "ten"))["a"]
    twelve = pptx_pages(write(tmp_path, [page("a", [element(), footer(12)])], "twelve"))["a"]
    thirteen = pptx_pages(write(tmp_path, [page("a", [element(), footer(13)])], "thirteen"))["a"]
    assert nine.choice != ten.choice and nine.full != ten.full
    assert (twelve.choice, twelve.full) == (thirteen.choice, thirteen.full)
    assert ten.footer == {"area": [660.0, 510.0, 36.0, 22.0], "text": "10"}


def test_notes_do_not_change_parts(tmp_path):
    pages = base_pages(tmp_path)
    changed = deepcopy(pages)
    changed[0]["notes"] = "別の説明をします。"
    changed[1]["notes"] = ""
    assert pptx_pages(write(tmp_path, pages, "first")) == pptx_pages(write(tmp_path, changed, "notes"))


def test_text_image_and_layout_changes_parts(tmp_path):
    pages = base_pages(tmp_path)
    first = pptx_pages(write(tmp_path, pages, "first"))
    text = deepcopy(pages)
    text[0]["elements"][0]["content"] = [{"kind": "text", "text": "別の文章"}]
    moved = deepcopy(pages)
    moved[0]["elements"][0]["area"] = [30, 60, 200, 70]
    image = deepcopy(pages)
    image[1]["elements"][0]["path"] = picture(tmp_path / "figure2.png", "green")
    for changed, name, slide_id, other in (
        (text, "text", "a", "b"),
        (moved, "moved", "a", "b"),
        (image, "image", "b", "a"),
    ):
        result = pptx_pages(write(tmp_path, changed, name))
        assert result[slide_id].choice != first[slide_id].choice, name
        assert result[slide_id].full != first[slide_id].full, name
        assert result[other] == first[other], name
    # 雛形のレイアウトを変えると全ページが変わる
    path = write(tmp_path, pages, "layout")
    parts = read_package(path)
    slide_name = slide_order(parts)[0][0]
    relations = etree.fromstring(parts[slide_name.replace("slides/", "slides/_rels/") + ".rels"])
    layout = (
        "ppt/slideLayouts/"
        + next(node.get("Target") for node in relations if node.get("Type").endswith("/slideLayout")).rsplit("/", 1)[-1]
    )
    root = etree.fromstring(parts[layout])
    root.find("p:cSld", NS).set("name", "Blank changed")
    parts[layout] = xml_bytes(root)
    write_package(parts, path)
    changed = pptx_pages(path)
    assert all(changed[slide_id].choice != first[slide_id].choice for slide_id in ("a", "b"))


def math_pdf(path, text):
    with Document() as document:
        document.new_page(width=720, height=540).insert_text((35, 60), text, fontsize=24)
        document.save(path)


def fallback_media(path):
    """mc:Fallback の画像が参照する部品名。"""
    parts = read_package(path)
    names = []
    for name in slide_order(parts)[0]:
        relations = {
            node.get("Id"): node.get("Target")
            for node in etree.fromstring(parts[name.replace("slides/", "slides/_rels/") + ".rels"])
        }
        for blip in etree.fromstring(parts[name]).findall(".//mc:Fallback//a:blip", NS):
            names.append("ppt/" + relations[blip.get(tag("r", "embed"))].removeprefix("../"))
    return names


def group_shape(source, output, element_id):
    """element_id の図形（mc:AlternateContent ならその全体）を、座標変換が恒等の p:grpSp に入れる。"""
    parts = read_package(source)
    name = slide_order(parts)[0][0]
    root = etree.fromstring(parts[name])
    tree = root.find("p:cSld/p:spTree", NS)
    target = next(
        child for child in tree if any(node.get("name") == element_id for node in child.iter(tag("p", "cNvPr")))
    )
    group = etree.Element(tag("p", "grpSp"))
    properties = etree.SubElement(group, tag("p", "nvGrpSpPr"))
    etree.SubElement(properties, tag("p", "cNvPr"), id="90", name="formula_group")
    etree.SubElement(properties, tag("p", "cNvGrpSpPr"))
    etree.SubElement(properties, tag("p", "nvPr"))
    transform = etree.SubElement(etree.SubElement(group, tag("p", "grpSpPr")), tag("a", "xfrm"))
    for local, values in (
        ("off", {"x": "381000", "y": "508000"}),
        ("ext", {"cx": "2540000", "cy": "889000"}),
        ("chOff", {"x": "381000", "y": "508000"}),
        ("chExt", {"cx": "2540000", "cy": "889000"}),
    ):
        etree.SubElement(transform, tag("a", local), **values)
    tree.replace(target, group)
    group.append(target)
    parts[name] = xml_bytes(root)
    write_package(parts, output)
    return output


def test_choice_ignores_fallback_and_matches_unwrapped(tmp_path):
    pages = [
        page(
            "m",
            [element("formula", content=[{"kind": "text", "text": "式 "}, {"kind": "math", "text": "x_i"}]), footer(1)],
        )
    ]
    native = write(tmp_path, pages, "native")
    math_pdf(tmp_path / "first.pdf", "x1")
    math_pdf(tmp_path / "second.pdf", "y2")
    final, other = tmp_path / "final.pptx", tmp_path / "other.pptx"
    finalize_math(native, tmp_path / "first.pdf", final)
    finalize_math(native, tmp_path / "second.pdf", other)
    unwrapped, wrapped, recropped = (pptx_pages(path)["m"] for path in (native, final, other))
    assert unwrapped.choice == wrapped.choice == recropped.choice
    assert len({unwrapped.full, wrapped.full, recropped.full}) == 3
    # 互換画像だけを仮の画像にした最終形の PPTX と、その変換から切り出した最終 PPTX は、スライドの XML と関係が同じで、
    # 互換画像の中身だけが違う。標準表示の証拠は同じなので、最終形の PPTX を変換した標準表示の確認は最終 PPTX の確認になる。
    # 互換表示（と LibreOffice の標準表示）の証拠は互換画像の中身を含むので、最終 PPTX を変換して確かめる。
    form, from_form = tmp_path / "final_form.pptx", tmp_path / "from_form.pptx"
    prepare_final_form(native, form)
    finalize_math(form, tmp_path / "first.pdf", from_form)
    assert compare_package_structure(form, from_form)["matched"]
    assert fallback_media(form) == fallback_media(from_form) != []
    temporary, finished = (pptx_pages(path)["m"] for path in (form, from_form))
    assert finished == wrapped
    assert temporary.choice == finished.choice and temporary.full != finished.full
    assert visual_digest(temporary, "powerpoint", "standard", ENVIRONMENT, render_date=DATE) == visual_digest(
        finished, "powerpoint", "standard", ENVIRONMENT, render_date=DATE
    )
    for renderer, variant in (("powerpoint", "fallback"), ("libreoffice", "standard")):
        assert visual_digest(temporary, renderer, variant, ENVIRONMENT, render_date=DATE, number=1) != visual_digest(
            finished, renderer, variant, ENVIRONMENT, render_date=DATE, number=1
        )
    # p:grpSp の中の mc:AlternateContent も同じ
    grouped_native = pptx_pages(group_shape(native, tmp_path / "grouped_native.pptx", "formula"))["m"]
    grouped_final = pptx_pages(group_shape(final, tmp_path / "grouped_final.pptx", "formula"))["m"]
    assert grouped_native.choice == grouped_final.choice != wrapped.choice
    assert grouped_native.full != grouped_final.full


def alternate(choice=None, fallback=None):
    node = etree.Element(tag("mc", "AlternateContent"), nsmap={"mc": NS["mc"]})
    for local, children in (("Choice", choice), ("Fallback", fallback)):
        if children is not None:
            branch = etree.SubElement(node, tag("mc", local), **({"Requires": "a14"} if local == "Choice" else {}))
            branch.extend(children)
    return node


def test_choice_uses_first_choice_or_fallback_of_nested_alternatives(tmp_path):
    pages = [
        page(
            "a",
            [
                element(name, area=(30, 40 + 80 * index, 200, 70), content=[{"kind": "text", "text": name}])
                for index, name in enumerate(("first", "second", "third", "fourth"))
            ],
        )
    ]
    original = write(tmp_path, pages, "original")
    parts = read_package(original)
    name = slide_order(parts)[0][0]
    root = etree.fromstring(parts[name])
    tree = root.find("p:cSld/p:spTree", NS)
    shapes = {node.find("p:nvSpPr/p:cNvPr", NS).get("name"): node for node in tree.findall("p:sp", NS)}
    first, second, third, fourth = (shapes[key] for key in ("first", "second", "third", "fourth"))
    # 互換表示の側に別の図形を置く。choice では最初の mc:Choice、なければ mc:Fallback を使い、入れ子も置き換える
    for target, build in (
        (first, lambda: alternate(choice=[first], fallback=[deepcopy(second)])),
        (second, lambda: alternate(fallback=[second])),
        (
            third,
            lambda: alternate(
                choice=[alternate(choice=[third], fallback=[deepcopy(first)]), fourth], fallback=[deepcopy(second)]
            ),
        ),
    ):
        position = tree.index(target)
        tree.insert(position, build())
    parts[name] = xml_bytes(root)
    write_package(parts, tmp_path / "alternatives.pptx")
    unwrapped, wrapped = pptx_pages(original)["a"], pptx_pages(tmp_path / "alternatives.pptx")["a"]
    assert wrapped.choice == unwrapped.choice == unwrapped.full
    assert wrapped.full != unwrapped.full


FIELD_SHAPE = (
    f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="90" name="Page number"/>'
    '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="8382000" y="6400800"/>'
    '<a:ext cx="609600" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
    "<a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>"
)


def test_fields_only_from_displayed_shapes(tmp_path):
    pages = [page("a", [element()]), page("b", [element()])]
    # 既定の雛形のマスターと空白レイアウトには、日付と番号のフィールドが配置用の図形の中だけにある
    Presentation().save(tmp_path / "plain.pptx")
    plain = pptx_pages(write(tmp_path, pages, "plain_deck", template={"path": "plain.pptx", "layout_index": 6}))
    assert plain["a"].fields == () and plain["b"].fields == ()
    # マスターの通常の図形のフィールドは含める
    template = Presentation()
    template.slide_master.element.find("p:cSld/p:spTree", NS).append(etree.fromstring(FIELD_SHAPE))
    template.save(tmp_path / "numbered.pptx")
    numbered = pptx_pages(
        write(
            tmp_path,
            pages,
            "numbered_deck",
            template={"path": "numbered.pptx", "layout_index": 6, "show_master_shapes": True},
        )
    )
    assert numbered["a"].fields == ("slidenum",) == numbered["b"].fields
    # フィールドがあると番号が証拠に入る。なければ番号によらない
    moved = replace(numbered["a"], number=2)
    assert visual_digest(numbered["a"], "powerpoint", "standard", ENVIRONMENT, render_date=DATE) != visual_digest(
        moved, "powerpoint", "standard", ENVIRONMENT, render_date=DATE
    )
    assert visual_digest(plain["a"], "powerpoint", "standard", ENVIRONMENT, render_date=DATE) == visual_digest(
        replace(plain["a"], number=2), "powerpoint", "standard", ENVIRONMENT, render_date=DATE
    )
    # 日付のフィールドがあれば生成した日も入る
    dated = replace(numbered["a"], fields=("datetime1", "slidenum"))
    assert visual_digest(dated, "powerpoint", "standard", ENVIRONMENT, render_date="2026-10-12") != visual_digest(
        dated, "powerpoint", "standard", ENVIRONMENT, render_date="2026-10-13"
    )
    assert visual_digest(
        numbered["a"], "powerpoint", "standard", ENVIRONMENT, render_date="2026-10-12"
    ) == visual_digest(numbered["a"], "powerpoint", "standard", ENVIRONMENT, render_date="2026-10-13")


def test_visual_digest_uses_choice_only_for_powerpoint_standard():
    parts = PageParts("a", 1, "c" * 64, "f" * 64, None, ())
    other_choice, other_full = replace(parts, choice="d" * 64), replace(parts, full="g" * 64)

    def digest(item, renderer, variant):
        return visual_digest(item, renderer, variant, ENVIRONMENT, render_date=DATE, number=1)

    assert digest(other_choice, "powerpoint", "standard") != digest(parts, "powerpoint", "standard")
    assert digest(other_full, "powerpoint", "standard") == digest(parts, "powerpoint", "standard")
    for renderer, variant in (("libreoffice", "standard"), ("powerpoint", "fallback")):
        assert digest(other_choice, renderer, variant) == digest(parts, renderer, variant)
        assert digest(other_full, renderer, variant) != digest(parts, renderer, variant)
    assert digest(parts, "powerpoint", "standard") != digest(parts, "powerpoint", "fallback")
    assert visual_digest(parts, "powerpoint", "standard", ENVIRONMENT, render_date=DATE) != visual_digest(
        parts, "powerpoint", "standard", "0" * 64, render_date=DATE
    )
    with pytest.raises(ValueError, match="表示方法"):
        visual_digest(parts, "powerpoint", "native", ENVIRONMENT, render_date=DATE)


def test_fallback_evidence_uses_position_in_clone():
    numbered = PageParts("formula", 5, "c" * 64, "f" * 64, None, ("slidenum",))
    # 確認用複製は数式のページだけなので、資料の中の位置ではなく複製の中の位置を明示する
    with pytest.raises(ValueError, match="確認用複製"):
        visual_digest(numbered, "powerpoint", "fallback", ENVIRONMENT, render_date=DATE)

    def fallback(item, number):
        return visual_digest(item, "powerpoint", "fallback", ENVIRONMENT, render_date=DATE, number=number)

    assert fallback(numbered, 2) == fallback(replace(numbered, number=7), 2)
    assert fallback(numbered, 2) != fallback(numbered, 3)
    # フィールドのないページは位置によらない
    plain = replace(numbered, fields=())
    assert fallback(plain, 2) == fallback(plain, 3)


def font_match(family, variant, table):
    return {
        "family": family,
        "variant": variant,
        "table_sha256": table,
        "windows_path": f"C:\\Windows\\Fonts\\{family}.ttf",
        "windows_index": 0,
        "source": f"/fonts/{family}.ttf",
        "font_names": [family],
        "version": ["1.0"],
        "index": 0,
    }


def test_display_environment_ignores_used_font_subset():
    detected = {
        "version": "16.0.1",
        "executable": "powershell.exe",
        "windows_user_sid": "S-1-5-21-1000",
        "state_dir": "C:\\Users\\editor\\AppData\\Local\\slidekit\\powerpoint",
        "installed_fonts": [{"name": "Arial", "path": "C:\\Windows\\Fonts\\arial.ttf"}],
        "font_matches": [font_match("Arial", "regular", "a1"), font_match("Arial", "bold", "a2")],
    }
    actual = {
        "renderer": "powerpoint",
        "version": "16.0.1",
        "build": "19000",
        "executable": "powershell.exe",
        "elapsed_seconds": 3.2,
        "input_sha256": "i1",
        "output_sha256": "o1",
    }
    # 実行環境で宣言したフォント。計測で使ったフォントの一覧（registry.inventory()）は引数に取らない
    declared = {"Arial": {"sha256": "f1", "index": 0, "bold_sha256": "f2", "bold_index": 0}}
    base = display_environment("powerpoint", detected, actual, declared, "1.28.2")
    unchanged = [
        # 登録された書体と状態のフォルダー（使う書体が増えると登録の一覧も増える）
        {
            **detected,
            "installed_fonts": detected["installed_fonts"]
            + [{"name": "Calibri", "path": "C:\\Windows\\Fonts\\calibri.ttf"}],
            "state_dir": "D:\\state",
        },
        {**detected, "font_matches": list(reversed(detected["font_matches"]))},
    ]
    for value in unchanged:
        assert display_environment("powerpoint", value, actual, declared, "1.28.2") == base
    assert (
        display_environment(
            "powerpoint",
            detected,
            {**actual, "elapsed_seconds": 9.0, "input_sha256": "i2", "output_sha256": "o2"},
            declared,
            "1.28.2",
        )
        == base
    )
    assert display_environment("powerpoint", detected, actual, declared, "1.28.2", scale=2) == base
    changed = [
        display_environment("powerpoint", detected, actual, {"Arial": {**declared["Arial"], "sha256": "f9"}}, "1.28.2"),
        display_environment(
            "powerpoint", detected, actual, {**declared, "Meiryo": {"sha256": "m1", "index": 0}}, "1.28.2"
        ),
        display_environment("powerpoint", {**detected, "version": "16.0.2"}, actual, declared, "1.28.2"),
        display_environment("powerpoint", detected, {**actual, "build": "19001"}, declared, "1.28.2"),
        display_environment(
            "powerpoint",
            {**detected, "font_matches": [font_match("Arial", "regular", "a9"), font_match("Arial", "bold", "a2")]},
            actual,
            declared,
            "1.28.2",
        ),
        display_environment("powerpoint", detected, actual, declared, "1.28.3"),
        display_environment("powerpoint", detected, actual, declared, "1.28.2", scale=4),
        display_environment("libreoffice", detected, actual, declared, "1.28.2"),
    ]
    assert base not in changed and len(set(changed)) == len(changed)


def synthetic_page(path, *, shift=0, drop=False, plus=False, number=(1, 2)):
    """2 倍で画像にした 720×540 pt のページを模した画像。文字は細い長方形の列で表す。"""
    image = Image.new("RGB", (1440, 1080), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, 1439, 124], fill="#183153")
    for line, y in enumerate((300, 360, 420)):
        if drop and line == 1:
            continue
        for index in range(24):
            x = 100 + index * 16 + (shift if line == 0 else 0)
            draw.rectangle([x, y, x + 8, y + 26], fill="#20252B")
    # 符号：横線（マイナス）と、縦線を加えたプラス
    draw.rectangle([100, 520, 140, 525], fill="#20252B")
    if plus:
        draw.rectangle([118, 503, 122, 542], fill="#20252B")
    # ページ番号（pt で [660, 510, 36, 22] の中）
    for index, digit in enumerate(number):
        x = 1330 + index * 14
        draw.rectangle([x, 1030, x + 8, 1036 + digit * 3], fill="#606B78")
    image.save(path, compress_level=1)
    return path


def test_render_difference_tolerates_one_pixel_shift(tmp_path):
    base = synthetic_page(tmp_path / "base.png")

    def strong(name, **changes):
        result = render_difference(base, synthetic_page(tmp_path / f"{name}.png", **changes))
        assert result["method"] == 1 and result["size_mismatch"] is False
        return result["strong_pixels"]

    assert strong("same") == 0
    assert strong("shift_one", shift=1) == 0
    for name, changes in (("shift_two", {"shift": 2}), ("dropped_line", {"drop": True}), ("sign", {"plus": True})):
        assert strong(name, **changes) > NOISE_PIXELS, name
    Image.new("RGB", (720, 540), "white").save(tmp_path / "small.png")
    assert render_difference(base, tmp_path / "small.png") == {
        "method": 1,
        "strong_pixels": None,
        "size_mismatch": True,
    }


def test_render_difference_masks_footer(tmp_path):
    base = synthetic_page(tmp_path / "base.png")
    renumbered = synthetic_page(tmp_path / "renumbered.png", number=(1, 3))
    assert render_difference(base, renumbered)["strong_pixels"] > 0
    assert render_difference(base, renumbered, masks=[[660.0, 510.0, 36.0, 22.0]])["strong_pixels"] == 0
    # 番号の領域の外の変化は、領域を除いても見つける
    changed = synthetic_page(tmp_path / "changed.png", number=(1, 3), drop=True)
    assert render_difference(base, changed, masks=[[660.0, 510.0, 36.0, 22.0]])["strong_pixels"] > NOISE_PIXELS
    with pytest.raises(ValueError, match="領域"):
        render_difference(base, renumbered, masks=[[660.0, 510.0, 36.0]])


def content_project(tmp_path, **changes):
    config = {
        "metadata": {"title": "資料", "author": "slidekit example"},
        "inputs": {
            "documents": [
                {"id": "survey", "path": str(tmp_path / "survey.md")},
                {"id": "manual", "path": str(tmp_path / "manual.md")},
            ],
            "bibliography": [],
        },
        "_bibliography_entries": {"textbook2001": "k1", "article2010": "l1"},
    }
    config.update(changes)
    return SimpleNamespace(
        config=config,
        rules={"notes_style": "unrestricted"},
        brief_path=tmp_path / "brief.md",
        input_files={
            str(tmp_path / "survey.md"): "s1",
            str(tmp_path / "manual.md"): "m1",
            str(tmp_path / "brief.md"): "b1",
        },
    )


def content_page():
    """作成 API の構築結果と同じ形のページ（画像は素材の準備の後）。"""
    style = {"size_pt": 20, "color": "20252B"}
    return {
        "id": "intro",
        "number": 3,
        "kind": "content",
        "title": "探索の手順",
        "notes": "手順を説明します。",
        "seconds": 60,
        "content_area": [36, 72, 648, 432],
        "allowed_warnings": [],
        "allow_roles": [],
        "elements": [
            {
                "id": "_theme/background",
                "kind": "box",
                "area": [0, 0, 720, 540],
                "role": "body",
                "content": [],
                "semantic": "background",
                "style": style,
            },
            {
                "id": "_theme/title",
                "kind": "text",
                "area": [27, 13, 666, 46],
                "role": "slide_title",
                "semantic": "title",
                "content": [{"kind": "text", "text": "探索の手順"}],
                "style": style,
            },
            {
                "id": "_theme/footer",
                "kind": "text",
                "area": [648, 512, 36, 20],
                "role": "note",
                "semantic": "footer",
                "content": [{"kind": "text", "text": "3"}],
                "style": style,
            },
            {
                "id": "lead",
                "kind": "text",
                "area": [36, 80, 600, 40],
                "role": "body",
                "style": style,
                "content": [
                    {"kind": "text", "text": "中央の値と比べる。"},
                    {"kind": "math", "text": "x_i"},
                    {
                        "kind": "text",
                        "text": "詳細",
                        "style": {"color": "183153", "underline": True},
                        "link": "https://example.com/guide",
                    },
                ],
                "sources": [{"document_id": "survey"}, "textbook2001"],
            },
            {
                "id": "figure",
                "kind": "image",
                "area": [36, 140, 200, 150],
                "actual_area": [36, 152, 200, 126],
                "role": "body",
                "style": style,
                "content": [],
                "path": "/run/0001/assets/abc.png",
                "source_path": "/deck/figure.svg",
                "source_sha256": "i1",
                "fit": "contain",
                "description": "探索の図",
            },
            {
                "id": "step",
                "kind": "box",
                "area": [300, 140, 120, 60],
                "role": "body",
                "shape": "rounded_rectangle",
                "style": style,
                "content": [{"kind": "text", "text": "比較"}],
            },
            {
                "id": "flow",
                "kind": "arrow",
                "area": [236, 170, 64, 0],
                "start": [236, 170],
                "end": [300, 170],
                "role": "body",
                "style": style,
                "content": [],
                "start_ref": {"element_id": "figure", "side": "right"},
                "end_ref": {"element_id": "step", "side": "left"},
            },
            {
                "id": "table",
                "kind": "group",
                "area": [36, 320, 400, 80],
                "role": "body",
                "style": style,
                "content": [],
                "table_format": "shapes",
                "row_keys": ["first"],
                "column_keys": ["name", "value"],
                "pptx_group": True,
            },
            {
                "id": "table/first/name",
                "kind": "box",
                "area": [36, 320, 200, 80],
                "role": "body",
                "style": style,
                "parent_id": "table",
                "shape": "rectangle",
                "content": [{"kind": "text", "text": "回数"}],
            },
        ],
    }


def item(slide, identifier):
    return next(element for element in slide["elements"] if element["id"] == identifier)


def test_content_digest_ignores_number_layout_style_and_time(tmp_path):
    current = content_project(tmp_path)
    base = content_digest(content_page(), current)

    def ignored(change):
        slide = content_page()
        change(slide)
        assert content_digest(slide, current) == base

    ignored(lambda slide: slide.update(number=12))
    ignored(lambda slide: slide.update(seconds=90, kind="references", content_area=[20, 60, 680, 460]))
    ignored(lambda slide: item(slide, "lead").update(area=[40, 90, 580, 50]))
    ignored(lambda slide: item(slide, "lead").update(style={"size_pt": 24, "color": "183153", "bold": True}))
    ignored(lambda slide: item(slide, "flow").update(area=[240, 180, 60, 0], start=[240, 180], end=[300, 180]))
    ignored(
        lambda slide: item(slide, "figure").update(
            actual_area=[36, 140, 200, 150],
            path="/run/0002/assets/abc.png",
            raster_scale=3,
            vector=True,
            vector_path="/run/0002/assets/abc.emf",
        )
    )
    ignored(lambda slide: slide["elements"].reverse())
    ignored(lambda slide: item(slide, "_theme/footer").update(content=[{"kind": "text", "text": "12"}]))
    ignored(
        lambda slide: slide.update(
            allowed_warnings=[{"code": "element_overlap", "element": "lead", "reason": "意図した配置"}],
            allow_roles=["note"],
        )
    )
    ignored(lambda slide: item(slide, "lead").update(allow_roles=["note"]))
    ignored(lambda slide: item(slide, "table").update(pptx_group=False))
    ignored(
        lambda slide: item(slide, "flow").update(
            start_ref={"element_id": "figure", "side": "right", "gap": 6, "at": 0.25}
        )
    )


def test_content_digest_tracks_text_notes_title_sources_and_structure(tmp_path):
    current = content_project(tmp_path)
    base = content_digest(content_page(), current)

    def tracked(change):
        slide = content_page()
        change(slide)
        assert content_digest(slide, current) != base

    tracked(lambda slide: item(slide, "lead")["content"][0].update(text="端の値と比べる。"))
    tracked(lambda slide: item(slide, "lead")["content"][1].update(text="x_j"))
    tracked(lambda slide: item(slide, "lead")["content"][0].update(style={"bold": True}))
    tracked(lambda slide: item(slide, "lead")["content"][2].update(link="https://example.com/other"))
    tracked(lambda slide: slide.update(notes="別の説明をします。"))
    tracked(lambda slide: slide.update(title="探索の考え方"))
    tracked(lambda slide: item(slide, "_theme/title")["content"][0].update(style={"color": "C43D3D"}))
    tracked(lambda slide: item(slide, "lead").update(sources=[{"document_id": "survey"}]))
    tracked(lambda slide: item(slide, "figure").update(source_sha256="i2"))
    tracked(lambda slide: item(slide, "figure").update(description="別の図"))
    tracked(lambda slide: item(slide, "flow").update(end_ref={"element_id": "lead", "side": "left"}))
    tracked(lambda slide: item(slide, "table").update(row_keys=["second"]))
    tracked(lambda slide: item(slide, "step").update(role="note"))
    tracked(
        lambda slide: slide["elements"].append(
            {
                "id": "extra",
                "kind": "text",
                "area": [0, 0, 10, 10],
                "role": "body",
                "content": [{"kind": "text", "text": "追加"}],
                "style": {},
            }
        )
    )
    tracked(lambda slide: slide["elements"].remove(item(slide, "step")))


def test_content_digest_uses_only_referenced_documents(tmp_path):
    current = content_project(tmp_path)
    base = content_digest(content_page(), current)

    def digest_with(**changes):
        changed = content_project(tmp_path)
        for name, value in changes.items():
            if name == "input_files":
                changed.input_files.update(value)
            elif name == "entries":
                changed.config["_bibliography_entries"].update(value)
            else:
                setattr(changed, name, value)
        return content_digest(content_page(), changed)

    assert digest_with(input_files={str(tmp_path / "survey.md"): "s2"}) != base
    assert digest_with(entries={"textbook2001": "k2"}) != base
    assert digest_with(input_files={str(tmp_path / "manual.md"): "m2"}) == base
    assert digest_with(entries={"article2010": "l2"}) == base
    assert digest_with(input_files={str(tmp_path / "brief.md"): "b2"}) == base
    assert digest_with(rules={"notes_style": "polite"}) == base
    changed = content_project(tmp_path, metadata={"title": "別の資料", "author": "slidekit example"})
    assert content_digest(content_page(), changed) == base
    # 出典の文字列は、元資料の識別子なら元資料、ほかは文献キーとして読む
    slide = content_page()
    item(slide, "lead")["sources"] = ["survey", {"bib_key": "textbook2001"}]
    written = content_digest(slide, current)
    changed = content_project(tmp_path)
    changed.input_files[str(tmp_path / "survey.md")] = "s2"
    assert content_digest(slide, changed) != written


def test_deck_digest_tracks_metadata_brief_and_order(tmp_path):
    current = content_project(tmp_path)
    pages = [content_page(), {**content_page(), "id": "summary", "title": "まとめ"}]
    deck = {"size": [720, 540], "slides": pages}
    base = deck_digest(deck, current)
    changed_metadata = content_project(tmp_path, metadata={"title": "別の資料", "author": "slidekit example"})
    changed_brief = content_project(tmp_path)
    changed_brief.input_files[str(tmp_path / "brief.md")] = "b2"
    without_brief = content_project(tmp_path)
    without_brief.brief_path = None
    changed = [
        deck_digest(deck, changed_metadata),
        deck_digest(deck, changed_brief),
        deck_digest(deck, without_brief),
        deck_digest({**deck, "slides": list(reversed(pages))}, current),
        deck_digest({**deck, "slides": [pages[0], {**pages[1], "title": "結論"}]}, current),
        deck_digest({**deck, "size": [960, 540]}, current),
    ]
    assert base not in changed and len(set(changed)) == len(changed)
    # ページの内容とノートは内容の確認で扱う
    edited = deepcopy(pages)
    edited[0]["notes"] = "別の説明をします。"
    item(edited[0], "lead")["content"] = [{"kind": "text", "text": "別の本文"}]
    assert deck_digest({**deck, "slides": edited}, current) == base
    # brief_path を持たない資料は依頼内容なしとして扱う
    legacy = SimpleNamespace(config=current.config, input_files=current.input_files)
    assert deck_digest(deck, legacy) == deck_digest(deck, without_brief)
    # 一覧画像の証拠は、一覧画像の画素・表示環境・資料全体の証拠で変わる
    overview = overview_digest(["c1", "c2"], ENVIRONMENT, base)
    assert overview == overview_digest(("c1", "c2"), ENVIRONMENT, base)
    assert overview not in {
        overview_digest(["c1"], ENVIRONMENT, base),
        overview_digest(["c1", "c2"], "0" * 64, base),
        overview_digest(["c1", "c2"], ENVIRONMENT, changed[0]),
    }


def scope(**changes):
    """diagnostics.EvidenceScope と同じ属性と既定値。"""
    values = {
        "rules": None,
        "roles": False,
        "fonts": True,
        "element": "all",
        "page": ("size", "notes"),
        "pair": False,
        "version": 1,
    }
    values.update(changes)
    return SimpleNamespace(**values)


def test_scoped_evidence_follows_scope():
    rules = {
        "forbidden_characters": [],
        "notes_style": "polite",
        "math_candidates": {"enabled": True, "variables": ["x"], "exclude_roles": ["note"]},
    }
    roles = {
        "body": {"size_pt": 20, "min_size_pt": 18},
        "note": {"size_pt": 14, "min_size_pt": 12},
        "heading": {"size_pt": 24, "min_size_pt": 20},
    }
    first = {
        "id": "lead",
        "kind": "text",
        "area": [36, 80, 300, 40],
        "role": "body",
        "style": {"size_pt": 20},
        "content": [{"kind": "text", "text": "本文"}],
    }
    second = {
        "id": "caption",
        "kind": "box",
        "area": [36, 100, 300, 40],
        "role": "note",
        "style": {"size_pt": 14},
        "content": [],
        "semantic": "decoration",
    }
    slide = {"id": "intro", "notes": "説明します。", "content_area": [36, 72, 648, 432]}

    def evidence(current, elements, **changes):
        return scoped_evidence(
            current,
            rules=changes.get("rules", rules),
            roles=roles,
            elements=elements,
            slide=changes.get("slide", slide),
            size=[720, 540],
        )

    # 既定は 0.1 と同じ範囲：記述規則の全体、要素の全体（識別子の順）、ページの寸法。ノートはページ単位の診断だけ
    assert evidence(scope(), [first, second]) == {"rules": rules, "elements": [second, first], "size": [720, 540]}
    assert evidence(scope(), []) == {"rules": rules, "elements": [], "size": [720, 540], "notes": "説明します。"}
    # 記述規則は指定した項目だけ
    limited = scope(rules=("math_candidates",), page=())
    assert evidence(limited, [first]) == {"rules": {"math_candidates": rules["math_candidates"]}, "elements": [first]}
    assert evidence(limited, [first], rules={**rules, "notes_style": "unrestricted"}) == evidence(limited, [first])
    assert evidence(scope(rules=(), page=()), [first])["rules"] == {}
    # 要素の内容だけ：位置・書式では変わらず、文字で変わる
    content = scope(rules=(), element="content", page=())
    assert evidence(content, [first])["elements"] == [
        {"kind": "text", "content": [{"kind": "text", "text": "本文"}], "role": "body", "semantic": None}
    ]
    assert evidence(content, [{**first, "area": [0, 0, 10, 10], "style": {"size_pt": 30}}]) == evidence(
        content, [first]
    )
    assert evidence(content, [{**first, "content": "別の本文"}]) != evidence(content, [first])
    # 役割の定義は、対象の要素の役割だけ
    assert evidence(scope(rules=(), roles=True, page=()), [first, second])["roles"] == {
        "body": roles["body"],
        "note": roles["note"],
    }
    # 本文領域
    area = scope(rules=(), page=("size", "content_area"))
    assert evidence(area, [first]) == {
        "rules": {},
        "elements": [first],
        "size": [720, 540],
        "content_area": [36, 72, 648, 432],
    }
    assert evidence(scope(rules=(), page=("notes",)), [], slide={**slide, "notes": "別の説明です。"}) != evidence(
        scope(rules=(), page=("notes",)), []
    )
    with pytest.raises(ValueError, match="element"):
        evidence(scope(element="position"), [first])
    with pytest.raises(ValueError, match="page"):
        evidence(scope(page=("title",)), [first])


def test_render_page_image_matches_previous_pixel_hash(tmp_path):
    pdf = tmp_path / "deck.pdf"
    with Document() as document:
        for text, color in (("First page", (0.1, 0.2, 0.3)), ("Second page", (0.8, 0.1, 0.1))):
            sheet = document.new_page(width=720, height=540)
            sheet.draw_rect((40, 40, 300, 120), color=color, fill=color)
            sheet.insert_text((60, 200), text, fontsize=28)
        document.save(pdf)
    hashes = []
    with open_pdf(pdf) as document:
        for index, sheet in enumerate(document):
            # 0.1 の pipeline._render_images と同じ計算（画像化は 2 倍）
            image = sheet.render(2)
            expected = hashlib.sha256(f"{image.width}x{image.height}:RGB:".encode() + image.tobytes()).hexdigest()
            thumbnail = image.copy()
            thumbnail.thumbnail((480, 270))
            output = tmp_path / "renders" / f"slide-{index + 1:03d}.png"
            result = render_page_image(str(pdf), index, str(output))
            assert result["path"] == str(output)
            assert result["pixel_hash"] == expected
            assert result["thumbnail_size"] == list(thumbnail.size) and result["thumbnail"] == thumbnail.tobytes()
            with Image.open(output) as saved:
                assert saved.size == (1440, 1080)
                assert pixel_digest(saved.width, saved.height, saved.convert("RGB").tobytes()) == expected
            hashes.append(result["pixel_hash"])
            # 子プロセスへの受け渡しに対応した値だけを返す
            assert pickle.loads(pickle.dumps(result)) == result
    assert len(set(hashes)) == 2
    plain = render_page_image(str(pdf), 1, str(tmp_path / "fallback.png"), thumbnail=False)
    assert plain == {"path": str(tmp_path / "fallback.png"), "pixel_hash": hashes[1]}
    assert pickle.loads(pickle.dumps(render_page_image)) is render_page_image
    assert Path(tmp_path / "fallback.png").is_file()
    with pytest.raises(ValueError, match="ページがない"):
        render_page_image(str(pdf), 2, str(tmp_path / "missing.png"))
    with pytest.raises(ValueError, match="画素の数"):
        pixel_digest(2, 2, bytes(11))
