"""元素材を変えずに変換し、指定外の参照を取得しないことを確かめる。"""

from hashlib import sha256
from pathlib import Path
from types import SimpleNamespace

from PIL import Image
import pytest

from pdf_tools import Document
from slidekit.assets import PDF_VECTOR_MESSAGE, prepare_assets


@pytest.fixture
def project(tmp_path):
    inputs = {}

    def resolve_input(value):
        path = Path(value).resolve()
        if not path.is_relative_to(tmp_path):
            raise ValueError("参照範囲外")
        inputs[str(path)] = sha256(path.read_bytes()).hexdigest()
        return path

    return SimpleNamespace(resolve_input=resolve_input, input_files=inputs)


def scene(path, *, page=None, fit="contain", area=(10, 20, 120, 100)):
    image = {"id": "picture", "kind": "image", "path": str(path), "fit": fit, "area": list(area), "content": []}
    if page is not None:
        image["page"] = page
    return {"slides": [{"id": "page", "elements": [image]}]}


def get_image(data):
    return data["slides"][0]["elements"][0]


@pytest.mark.parametrize("format,suffix", [("PNG", ".png"), ("JPEG", ".jpg")])
def test_raster_input_remains_byte_identical(project, tmp_path, format, suffix):
    source = tmp_path / ("source" + suffix)
    Image.new("RGB", (400, 200), "white").save(source, format=format)
    original = source.read_bytes()
    data = scene(source)
    records = prepare_assets(data, project, tmp_path / "converted")
    image = get_image(data)
    assert source.read_bytes() == original
    assert image["path"] == str(source)
    assert image["source_sha256"] == sha256(original).hexdigest()
    assert records[0]["image_size"] == [400, 200]
    assert image["actual_area"] == [10, 40, 120, 60]


def make_pdf(path):
    with Document() as document:
        for color in ((1, 0, 0), (0, 0, 1)):
            page = document.new_page(width=200, height=100)
            page.draw_rect(page.rect, color=color, fill=color)
        document.save(path)


def test_pdf_page_is_explicit_and_selected_content_is_rasterized(project, tmp_path):
    source = tmp_path / "original.pdf"
    make_pdf(source)
    original = source.read_bytes()
    data = scene(source, page=2)
    records = prepare_assets(data, project, tmp_path / "converted")
    image = get_image(data)
    with Image.open(image["path"]) as output:
        pixel = output.getpixel((output.width // 2, output.height // 2))
        assert pixel == (0, 0, 255)
        assert output.width / output.height == pytest.approx(2)
    assert source.read_bytes() == original
    assert records[0]["page"] == 2
    assert image["actual_area"] == [10, 40, 120, 60]


@pytest.mark.parametrize("page", [None, 0, True, 3])
def test_pdf_rejects_unspecified_invalid_or_absent_page(project, tmp_path, page):
    source = tmp_path / "original.pdf"
    make_pdf(source)
    with pytest.raises(ValueError, match="ページ|page"):
        prepare_assets(scene(source, page=page), project, tmp_path / "converted")


def test_cover_preserves_full_frame_and_source_ratio(project, tmp_path):
    source = tmp_path / "original.pdf"
    make_pdf(source)
    data = scene(source, page=1, fit="cover")
    prepare_assets(data, project, tmp_path / "converted")
    image = get_image(data)
    assert image["actual_area"] == image["area"]
    assert image["image_size"][0] / image["image_size"][1] == pytest.approx(2)


def test_self_contained_svg_keeps_ratio_and_original_bytes(project, tmp_path):
    source = tmp_path / "original.svg"
    source.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        '<defs><linearGradient id="g"><stop stop-color="red"/></linearGradient></defs>'
        '<rect width="200" height="100" fill="url(#g)"/></svg>'
    )
    original = source.read_bytes()
    data = scene(source)
    prepare_assets(data, project, tmp_path / "converted")
    image = get_image(data)
    assert source.read_bytes() == original
    assert image["image_size"][0] / image["image_size"][1] == pytest.approx(2)
    assert image["actual_area"] == [10, 40, 120, 60]


def test_svg_cover_has_enough_pixels_for_height(project, tmp_path):
    source = tmp_path / "wide.svg"
    source.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="10">'
        '<rect width="1000" height="10" fill="red"/></svg>'
    )
    data = scene(source, fit="cover", area=(0, 0, 50, 100))
    prepare_assets(data, project, tmp_path / "converted")
    width, height = get_image(data)["image_size"]
    assert height >= 200
    assert width / height == pytest.approx(100)


def test_svg_rejects_css_escaped_external_reference(project, tmp_path):
    source = tmp_path / "external.svg"
    source.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
        r'<rect width="100" height="100" style="fill:u\72l(https://example.invalid/a.svg#g)"/>'
        "</svg>"
    )
    with pytest.raises(ValueError, match="参照素材"):
        prepare_assets(scene(source), project, tmp_path / "converted")


@pytest.mark.parametrize(
    "body",
    [
        '<image href="https://example.invalid/image.png"/>',
        '<image href="file:///tmp/private.png"/>',
        '<style>@import "https://example.invalid/style.css";</style>',
        '<rect style="fill:url(https://example.invalid/image.svg#g)"/>',
        "<script>bad()</script>",
        '<foreignObject width="10" height="10"><body>html</body></foreignObject>',
    ],
)
def test_svg_external_references_and_active_content_fail_before_conversion(project, tmp_path, monkeypatch, body):
    import cairosvg

    calls = []

    def unexpected_conversion(*args, **kwargs):
        calls.append(kwargs)
        pytest.fail("不正な SVG を変換器へ渡している")

    # 本体が呼ぶ変換の関数を差し替える（変換中の取得関数が後で拒否する場合も、変換の前に拒否することを確かめる）。
    monkeypatch.setattr(cairosvg.surface.PNGSurface, "convert", unexpected_conversion)
    source = tmp_path / "external.svg"
    source.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">' + body + "</svg>")
    original = source.read_bytes()
    with pytest.raises(ValueError, match="SVG"):
        prepare_assets(scene(source), project, tmp_path / "converted")
    assert source.read_bytes() == original and calls == []


def test_svg_conversion_uses_patched_converter(project, tmp_path, monkeypatch):
    import cairosvg

    # 上の試験が差し替える関数を、本体が実際に呼ぶこと（自己完結した SVG は変換器へ渡る）。
    calls = []

    def conversion(*args, **kwargs):
        calls.append(sorted(kwargs))
        raise RuntimeError("変換器を呼んだ")

    monkeypatch.setattr(cairosvg.surface.PNGSurface, "convert", conversion)
    valid = tmp_path / "valid.svg"
    valid.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="10" height="10"/></svg>'
    )
    with pytest.raises(RuntimeError, match="変換器を呼んだ"):
        prepare_assets(scene(valid), project, tmp_path / "converted")
    assert calls == [["bytestring", "output_width", "url_fetcher"]]


def test_svg_external_entity_is_rejected(project, tmp_path):
    source = tmp_path / "entity.svg"
    source.write_text(
        '<!DOCTYPE svg [<!ENTITY private SYSTEM "file:///tmp/private">]>'
        '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><text>&private;</text></svg>'
    )
    with pytest.raises(ValueError, match="外部実体"):
        prepare_assets(scene(source), project, tmp_path / "converted")


@pytest.mark.parametrize("encoding", ["utf-16", "utf-32"])
def test_svg_doctype_in_other_encoding_is_rejected(project, tmp_path, encoding):
    # バイト列の検索に掛からない文字コードでも、解析した文書の文書型宣言で拒否する。
    source = tmp_path / "entity.svg"
    source.write_bytes(
        (
            '<!DOCTYPE svg [<!ENTITY private SYSTEM "file:///tmp/private">]>'
            '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><text>&private;</text></svg>'
        ).encode(encoding)
    )
    data = scene(source)
    get_image(data)["vector"] = True
    with pytest.raises(ValueError, match="外部実体"):
        prepare_assets(data, project, tmp_path / "converted")
    assert not (tmp_path / "converted").exists() or not any((tmp_path / "converted").iterdir())


def test_svg_with_large_embedded_image_is_accepted():
    from slidekit.assets import _check_svg

    # 1 つの data: 属性が libxml2 の既定の上限（10,000,000 文字）を超える自己完結した SVG も受け付ける。
    payload = "A" * 10_000_004
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="10" height="10">'
        f'<image width="10" height="10" xlink:href="data:image/png;base64,{payload}"/></svg>'
    ).encode()
    _check_svg(svg)


def test_deck_vector_setting_applies_only_to_svg(project, tmp_path):
    # 資料の既定の vector は SVG の素材だけに適用し、PDF の素材は PNG だけにする（誤りにしない）。
    source = tmp_path / "original.pdf"
    make_pdf(source)
    svg = tmp_path / "wide.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><rect width="200" height="100"/></svg>'
    )
    pdf_image, svg_image = {**get_image(scene(source, page=1)), "id": "pdf"}, {**get_image(scene(svg)), "id": "svg"}
    data = {"slides": [{"id": "page", "elements": [pdf_image, svg_image]}]}
    pdf_record, svg_record = prepare_assets(data, configured(project, vector=True), tmp_path / "converted")
    assert (pdf_record["vector_output"], pdf_record["vector_sha256"]) == (None, None)
    assert "vector_path" not in pdf_image and Path(pdf_image["path"]).suffix == ".png"
    with Image.open(pdf_image["path"]) as output:
        assert output.format == "PNG" and output.size == tuple(pdf_record["image_size"])
    assert svg_record["vector_output"] == svg_image["vector_path"]
    assert Path(svg_image["vector_path"]).read_bytes() == svg.read_bytes()


@pytest.mark.parametrize(
    "text",
    [
        '<svg xmlns="http://www.w3.org/2000/svg"><rect></svg>',
        '<svg xmlns="http://www.w3.org/2000/svg">',
        "not xml",
    ],
)
def test_malformed_svg_is_value_error(project, tmp_path, text):
    from slidekit.assets import _check_svg

    with pytest.raises(ValueError, match="SVG を XML として解析できない"):
        _check_svg(text.encode())
    # 素材の準備でも、入力の誤り（ValueError）として変換の前に止まる。
    source = tmp_path / "broken.svg"
    source.write_text(text)
    with pytest.raises(ValueError, match="SVG を XML として解析できない"):
        prepare_assets(scene(source), project, tmp_path / "converted")


def test_corrupt_raster_and_unknown_format_are_not_accepted(project, tmp_path):
    source = tmp_path / "corrupt.png"
    source.write_bytes(b"This is not a PNG")
    with pytest.raises(OSError):
        prepare_assets(scene(source), project, tmp_path / "converted")
    unknown = tmp_path / "image.bmp"
    Image.new("RGB", (10, 10)).save(unknown)
    with pytest.raises(ValueError, match="未対応"):
        prepare_assets(scene(unknown), project, tmp_path / "converted")


def test_assets_use_hash_recorded_by_input_resolution(project, tmp_path, monkeypatch):
    import slidekit.assets as assets

    # 素材のハッシュ値は resolve_input が記録した値を使い、変換で読み直さない。
    def unexpected_hash(path):
        pytest.fail(f"素材をハッシュし直している: {path}")

    monkeypatch.setattr(assets, "file_hash", unexpected_hash)
    source = tmp_path / "original.pdf"
    make_pdf(source)
    data = scene(source, page=1)
    records = prepare_assets(data, project, tmp_path / "converted")
    assert records[0]["sha256"] == get_image(data)["source_sha256"] == project.input_files[str(source)]


def configured(project, **images):
    project.config = {"images": {"raster_scale": 2, "vector": False, **images}}
    return project


def test_raster_scale_from_deck_and_image(project, tmp_path):
    source = tmp_path / "original.pdf"
    make_pdf(source)
    # 既定（資料の設定がない簡易な資料を含む）は配置寸法 1 pt あたり 2 画素で、 0.1 と同じ名前と画素数になる。
    data = scene(source, page=1, area=(0, 0, 100, 50))
    record = prepare_assets(data, project, tmp_path / "default")[0]
    assert (record["target_pixels"], record["scale"], get_image(data)["image_size"]) == ([200, 100], 2, [200, 100])
    assert (record["vector_output"], record["vector_sha256"]) == (None, None) and "vector_path" not in get_image(data)
    # deck.toml の [images] raster_scale を使い、画像の指定がそれに優先する。
    data = scene(source, page=1, area=(0, 0, 100, 50))
    record = prepare_assets(data, configured(project, raster_scale=3), tmp_path / "deck")[0]
    assert (record["target_pixels"], record["scale"], get_image(data)["image_size"]) == ([300, 150], 3, [300, 150])
    data = scene(source, page=1, area=(0, 0, 100, 50))
    get_image(data)["raster_scale"] = 1.5
    record = prepare_assets(data, project, tmp_path / "image")[0]
    assert (record["target_pixels"], record["scale"]) == ([150, 75], 1.5)
    # SVG も同じ規則で画素数を決める。
    svg = tmp_path / "wide.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><rect width="200" height="100"/></svg>'
    )
    data = scene(svg, area=(0, 0, 100, 50))
    get_image(data)["raster_scale"] = 4
    assert prepare_assets(data, project, tmp_path / "svg")[0]["target_pixels"] == [400, 200]
    assert get_image(data)["image_size"] == [400, 200]


def test_raster_options_are_rejected_for_png_and_out_of_range(project, tmp_path):
    from slidekit.errors import InputError

    png = tmp_path / "picture.png"
    Image.new("RGB", (40, 20), "white").save(png)
    for name, value in (("raster_scale", 2), ("vector", False)):
        data = scene(png)
        get_image(data)[name] = value
        with pytest.raises(InputError, match="PDF・SVG の素材だけ") as raised:
            prepare_assets(data, project, tmp_path / "converted")
        assert (raised.value.code, raised.value.element_id) == ("invalid_image_option", "picture")
    source = tmp_path / "original.pdf"
    make_pdf(source)
    for value in (0.5, 17, True, float("nan"), "2"):
        data = scene(source, page=1)
        get_image(data)["raster_scale"] = value
        with pytest.raises(InputError, match="raster_scale は 1 以上 16 以下"):
            prepare_assets(data, project, tmp_path / "converted")
    data = scene(source, page=1)
    get_image(data)["vector"] = "yes"
    with pytest.raises(InputError, match="vector は真偽値"):
        prepare_assets(data, project, tmp_path / "converted")
    # 範囲の端は受け付ける。
    for value in (1, 16):
        data = scene(source, page=1, area=(0, 0, 10, 5))
        get_image(data)["raster_scale"] = value
        assert prepare_assets(data, project, tmp_path / f"edge_{value}")[0]["target_pixels"] == [10 * value, 5 * value]


def test_pdf_vector_is_rejected(project, tmp_path):
    from slidekit.errors import InputError

    # PDF の素材は PNG だけにする。画像に vector を指定した PDF の素材は、変換の前に誤りにする。
    source = tmp_path / "figure.pdf"
    make_pdf(source)
    original = source.read_bytes()
    for settings in ({}, {"vector": True}):
        data = scene(source, page=1)
        get_image(data)["vector"] = True
        with pytest.raises(InputError) as raised:
            prepare_assets(data, configured(project, **settings), tmp_path / "converted")
        assert PDF_VECTOR_MESSAGE in str(raised.value)
        assert (raised.value.code, raised.value.element_id) == ("invalid_image_option", "picture")
        assert "vector_path" not in get_image(data) and get_image(data)["path"] == str(source)
    assert source.read_bytes() == original
    assert not (tmp_path / "converted").exists()
    # vector に偽を指定した PDF の素材は、これまでどおり PNG にする。
    data = scene(source, page=1)
    get_image(data)["vector"] = False
    record = prepare_assets(data, project, tmp_path / "converted")[0]
    assert (record["vector_output"], record["vector_sha256"]) == (None, None) and "vector_path" not in get_image(data)
    with Image.open(get_image(data)["path"]) as output:
        assert output.format == "PNG" and output.size == tuple(record["image_size"])


def test_svg_vector_keeps_checked_source_bytes(project, tmp_path):
    source = tmp_path / "original.svg"
    source.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><rect width="200" height="100" fill="red"/></svg>'
    )
    data = scene(source)
    get_image(data)["vector"] = True
    record = prepare_assets(data, project, tmp_path / "converted")[0]
    assert Path(get_image(data)["vector_path"]).read_bytes() == source.read_bytes()
    assert record["vector_output"] == get_image(data)["vector_path"]
    # 外部を参照する SVG は、SVG でも埋め込む場合も変換の前に拒否する。
    external = tmp_path / "external.svg"
    external.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><image href="https://example.invalid/a.png"/></svg>'
    )
    data = scene(external)
    get_image(data)["vector"] = True
    with pytest.raises(ValueError, match="参照素材"):
        prepare_assets(data, project, tmp_path / "rejected")
    assert not (tmp_path / "rejected").exists() or not any((tmp_path / "rejected").glob("*.svg"))
