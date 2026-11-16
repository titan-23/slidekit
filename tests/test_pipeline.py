"""実 PPTX と PDF を使い、構築から確認・反映までを結合して確かめる。

確認記録の規則そのもの（証拠の一致、入力・成果物の変更、取り消し）は `test_records.py` が合成した実行で確かめる。
ここでは、生成の流れ（配置だけの検査、事前確認、部分生成・下書き、選んだ章だけの構築、変換の回数、互換画像、停止と失敗、記録）、
生成が書く記録（形式 2）と証拠が実際の成果物と一致すること、生成の後の確認の自動の引き継ぎ、確認から反映までが通ることを確かめる。
表示ソフトはフィクスチャ fake_session の偽の変換セッションに差し替える（保存した PPTX から検査用の PDF を作る）。
資料コードの読み込み（補助モジュールの追跡と再読み込み）は、生成をせずに `construct` で確かめる。
"""

import io
import json
import os
from pathlib import Path
import sys
import tomllib

from lxml import etree
from PIL import Image
from pptx import Presentation
import pytest

from pdf_tools import Document
from slidekit import fallback, fonts, pipeline, records
from slidekit.config import load_project
from slidekit.evidence import (
    EVIDENCE_VERSION,
    content_digest,
    deck_digest,
    display_environment,
    overview_digest,
    pixel_digest,
    pptx_pages,
    visual_digest,
)
from slidekit.files import file_hash
from slidekit.model import Deck
from slidekit.ooxml import read_package, write_package
from slidekit.pdf import open_pdf
from slidekit.renderers import RendererError
from slidekit.validators import make_diagnostic

# 試験の資料のページには日付のフィールドがないので、表示の証拠は生成した日によらない。
DATE = "2026-10-12"

DECK = """schema_version = 1
entrypoint = "slides.py:build"
size = "16:9"
[metadata]
title = "結合検査"
author = ""
last_modified_by = ""
[inputs]
read_roots = ["."]
[deliverables]
pptx = "presentation.pptx"
pdf = "presentation.pdf"
notes = "notes.md"
record = "deliverables.json"
"""

SLIDES = """from slidekit import Deck, Rect

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="intro", kind="blank", title="Summary", notes="確認内容を説明します。", seconds=45)
    slide.text(id="body", text="Process checked", area=Rect(60, 110, 800, 90))
    return deck
"""

# 枠に収まらない本文（配置のエラー text_overflow になる）。
LONG = "Process checked " * 20


def pages_source(*pages, math=(), overflow=(), extra=None, titles=None):
    """content 種別（ページ番号の飾りがある）のページを並べた生成コード。

    math のページには数式の文字枠を、overflow のページには枠に収まらない本文を置く。extra はページの識別子から追加の行への対応。
    題名は titles（ページの識別子から題名）か、なければ「Page 番号」にする。
    """
    lines = ["from slidekit import Deck, Math, Rect", "", "def build(context):", "    deck = Deck(context)"]
    for number, page in enumerate(pages, 1):
        title = (titles or {}).get(page, f"Page {number}")
        lines.append(
            f'    slide = deck.slide(id="{page}", kind="content", title="{title}", notes="説明します。", seconds=30)'
        )
        text = LONG if page in overflow else "Process checked"
        lines.append(f'    slide.text(id="body", text="{text}", area=Rect(60, 110, 400, 60))')
        if page in math:
            lines.append('    slide.text(id="equation", text=[Math("x_i^2")], area=Rect(60, 200, 300, 60))')
        lines.extend((extra or {}).get(page, ()))
    lines.append("    return deck")
    return "\n".join(lines) + "\n"


@pytest.fixture
def local_project(tmp_path, write_deck, dejavu, dejavu_environment, fake_session):
    """DejaVu Sans だけを使う 1 ページの資料。表示ソフトは偽の変換セッション（保存した PPTX から検査用の PDF を作る）にする。"""
    return write_deck(tmp_path / "deck", deck=DECK, slides=SLIDES, theme=dejavu.theme)


def rewrite(root, source, *, math=False, settings=""):
    """生成コードを書き換え、math が真なら数式の互換画像を必須にする。settings は deck.toml の末尾に加える。"""
    (root / "slides.py").write_text(source, encoding="utf-8")
    if math:
        settings = '[math]\nfallback = "required"\n' + settings
    if settings:
        path = root / "deck.toml"
        path.write_text(path.read_text(encoding="utf-8") + settings, encoding="utf-8")
    return root


def built(root, **options):
    result = pipeline.build(root / "deck.toml", **options)
    assert result["exit_code"] == 0, result
    return Path(result["build"])


def manifest(run):
    """生成記録をそのまま読む（記録の検査は records が行う）。"""
    return json.loads((run / "manifest.json").read_text(encoding="utf-8"))


def finish_review(run):
    records.record_review(run, "visual", "all", "検査担当", "各ページの表示を確認")
    records.record_review(run, "overview", "all", "検査担当", "ページ順と一覧を確認")
    records.record_review(run, "content", "all", "検査担当", "本文とノートを照合")
    # 資料コードで宣言した警告（declared）は許容済みなので、残りだけを許容する。
    for warning in records.check_build(run)["warnings"]:
        if not warning.get("declared"):
            records.accept_warning(
                run, warning["diagnostic_id"], "検査担当", "結合検査用の表示条件に起因することを確かめた"
            )


def pixels(path):
    with Image.open(path) as image:
        image = image.convert("RGB")
        return pixel_digest(image.width, image.height, image.tobytes())


def codes(result, severity=None):
    return sorted(issue["code"] for issue in result["issues"] if severity is None or issue["severity"] == severity)


OOXML = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def add_math_page(root):
    """数式のある 2 ページ目を加え、数式の互換画像を必須にする。"""
    settings = root / "deck.toml"
    settings.write_text(settings.read_text(encoding="utf-8") + '[math]\nfallback = "required"\n', encoding="utf-8")
    source = root / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8")
        .replace("from slidekit import Deck, Rect", "from slidekit import Deck, Math, Rect")
        .replace(
            "    return deck",
            """    formula = deck.slide(id="formula", kind="blank", title="Formula", notes="式を説明します。", seconds=30)
    formula.text(id="equation", text=[Math("x_i^2")], area=Rect(60, 110, 300, 60))
    return deck""",
        ),
        encoding="utf-8",
    )


# ---------------------------------------------------------------- 配置だけの検査と事前確認


def test_layout_runs_without_renderer_or_writes(local_project, monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("配置だけの検査は表示ソフトを使わない")

    monkeypatch.setattr(pipeline, "open_session", refuse)
    before = sorted(local_project.rglob("*"))
    result = pipeline.layout(local_project)
    assert (result["command"], result["status"], result["exit_code"]) == ("layout", "passed", 0)
    assert result["slides"] == 1 and result["selection"] is None and result["fallback_required"] is False
    # 予測は layout.json と同じ形で返す（ファイルは書かない）。
    (page,) = result["layout"]["slides"]
    assert page["id"] == "intro" and [element["id"] for element in page["elements"]] == ["body"]
    assert [item["stage"] for item in result["timings"]] == ["construct", "assets", "layout"]
    assert sorted(local_project.rglob("*")) == before and not (local_project / "out").exists()


def test_layout_limits_report_to_selected_slides(local_project):
    rewrite(local_project, pages_source("first", "second", "third", overflow=("second",)))
    second = pipeline.layout(local_project, slides="2")
    assert (second["status"], second["exit_code"]) == ("failed", 1)
    assert [
        (issue["code"], issue["slide_id"], issue["slide_number"])
        for issue in second["issues"]
        if issue["severity"] == "error"
    ] == [("text_overflow", "second", 2)]
    assert second["selection"] == {
        "spec": "2",
        "slide_ids": ["second"],
        "numbers": [2],
        "total_slides": 3,
        "construct": "all",
    }
    assert [page["id"] for page in second["layout"]["slides"]] == ["second"]
    first = pipeline.layout(local_project, slides="1")
    assert (first["status"], first["exit_code"], codes(first, "error")) == ("passed", 0, [])
    # 書式の誤りは例外にする（CLI が終了コード 2 にする）。
    with pytest.raises(ValueError, match="--slides"):
        pipeline.layout(local_project, slides="3-")


def test_doctor_keeps_layout_diagnostics_when_preflight_fails(local_project, fake_session):
    rewrite(local_project, pages_source("first", "second", overflow=("second",)))

    def busy(on_event):
        raise RendererError("SLIDEKIT_BUSY: 同じ Windows 利用者による変換が実行中です", code="renderer_busy")

    fake_session.on_open = busy
    result = pipeline.doctor(local_project)
    assert (result["status"], result["exit_code"]) == ("failed", 2)
    assert {"text_overflow", "renderer_busy"} <= set(codes(result))
    busy_issue = next(issue for issue in result["issues"] if issue["code"] == "renderer_busy")
    assert busy_issue["scope"] == "doctor"
    assert result["powerpoint"]["status"] == "failed"


@pytest.mark.parametrize("math_page", [False, True])
def test_doctor_reports_environment_without_writing_to_project(local_project, math_page):
    if math_page:
        add_math_page(local_project)
    before = sorted(local_project.rglob("*"))
    result = pipeline.doctor(local_project)
    assert result["status"] == "passed" and result["exit_code"] == 0
    assert result["slides"] == (2 if math_page else 1)
    assert result["fallback_required"] is math_page
    assert not [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert result["environment"]["fonts"]["fonts"]["DejaVu Sans"]["sha256"]
    assert result["environment"]["renderers"]["powerpoint"]["version"] == "fixture-1"
    assert result["powerpoint"] == {"status": "passed", "version": "fixture-1", "message": None}
    assert sorted(local_project.rglob("*")) == before


def test_doctor_rejects_unwritable_output_before_rendering(local_project, monkeypatch):
    import tempfile

    def denied(*args, **kwargs):
        raise PermissionError("read only")

    monkeypatch.setattr(tempfile, "TemporaryFile", denied)
    with pytest.raises(OSError, match="出力先へ書き込めません"):
        pipeline.doctor(local_project)


def test_doctor_reports_font_directories(local_project, dejavu, dejavu_environment, tmp_path):
    folder = tmp_path / "fonts"
    folder.mkdir()
    head, rest = dejavu.environment.split("\n", 1)
    dejavu_environment.write_text(f"{head}\nfont_directories = [{json.dumps(str(folder))}]\n{rest}", encoding="utf-8")
    result = pipeline.doctor(local_project)
    assert result["exit_code"] == 0 and result["environment"]["font_directories"] == [str(folder)]
    # 既定フォントのフォルダーは確認の証拠に入れないので、生成記録の表示の証拠はフォルダーによらない。
    run = built(local_project)
    record = manifest(run)
    assert record["environment"]["font_directories"] == [str(folder)]
    dejavu_environment.write_text(dejavu.environment, encoding="utf-8")
    other = manifest(built(local_project))
    assert other["environment"]["font_directories"] == []
    assert other["slides"][0]["visuals"] == record["slides"][0]["visuals"]


def test_stopped_build_records_font_directories(local_project, dejavu, dejavu_environment, tmp_path, fake_session):
    # 変換の前に止めた実行でも、既定フォントのフォルダー（設定の値か、設定がなければ None）を生成記録に入れる。
    rewrite(local_project, pages_source("first", overflow=("first",)))
    folder = tmp_path / "fonts"
    folder.mkdir()
    head, rest = dejavu.environment.split("\n", 1)
    dejavu_environment.write_text(f"{head}\nfont_directories = [{json.dumps(str(folder))}]\n{rest}", encoding="utf-8")
    stopped = pipeline.build(local_project / "deck.toml", stop_on_error=True)
    assert stopped["status"] == "stopped" and fake_session.sessions == []
    assert manifest(Path(stopped["build"]))["environment"]["font_directories"] == [str(folder)]
    dejavu_environment.write_text(dejavu.environment, encoding="utf-8")
    other = pipeline.build(local_project / "deck.toml", stop_on_error=True)
    environment = manifest(Path(other["build"]))["environment"]
    assert "font_directories" in environment and environment["font_directories"] is None


def test_doctor_detects_fonts_only_after_passed_preflight(local_project, fake_session, monkeypatch):
    # 事前確認に失敗したセッションは Windows への問い合わせを記憶していないことがあるので、自動探索をしない
    # （WSL では探索が PowerShell を新たに起動する）。
    detected = []
    monkeypatch.setattr(fonts, "_detect_font_directories", lambda environment: detected.append(True) or ())

    def busy(on_event):
        raise RendererError("SLIDEKIT_BUSY: 同じ Windows 利用者による変換が実行中です", code="renderer_busy")

    fake_session.on_open = busy
    failed = pipeline.doctor(local_project)
    assert failed["powerpoint"]["status"] == "failed"
    assert failed["environment"]["font_directories"] is None and detected == []
    fake_session.on_open = None
    passed = pipeline.doctor(local_project)
    assert passed["powerpoint"]["status"] == "passed"
    assert passed["environment"]["font_directories"] == [] and detected == [True]


def test_cli_keeps_machine_readable_stdout(local_project, monkeypatch, capsys):
    # 資料コードなどの print は標準エラーへ回し、--json の標準出力は結果の JSON だけにする。
    from slidekit.cli import main

    def noisy(path, **options):
        print("dependency notice")
        return {"command": "doctor", "status": "passed", "exit_code": 0}

    monkeypatch.setattr(pipeline, "doctor", noisy)
    assert main(["doctor", "--project", str(local_project), "--json"]) == 0
    captured = capsys.readouterr()
    assert json.loads(captured.out)["status"] == "passed"
    assert "dependency notice" in captured.err


# ---------------------------------------------------------------- 生成の流れと変換の回数


def test_build_uses_one_session_for_all_powerpoint_renders(local_project, fake_session, monkeypatch):
    add_math_page(local_project)
    hashed = []
    monkeypatch.setattr(fallback, "file_hash", lambda path: hashed.append(Path(path).name) or file_hash(path))
    run = built(local_project)
    # 互換画像の 3 つの処理は入力のハッシュ値を書き出しと変換で求めた値で受け取り、自分の出力だけを読む。
    assert sorted(hashed) == ["deck.pptx", "fallback.pptx", "final_form.pptx"]
    # PowerPoint の変換は 1 つのセッションで、標準表示 1 回と互換表示 1 回だけ（中間 PPTX は変換しない）。
    assert len(fake_session.sessions) == fake_session.doctors == fake_session.closed == 1
    assert [render["label"] for render in fake_session.renders] == ["standard", "fallback"]
    holder = fake_session.sessions[0]["holder"]
    assert (holder["command"], holder["run"], holder["project"]) == (
        "build",
        str(run),
        str(local_project / "deck.toml"),
    )
    record = manifest(run)
    operations = record["operations"]
    assert operations["sessions"]["powerpoint"]["renders"] == 2
    assert not (run / "intermediate/native.pdf").exists()
    # 標準表示は最終形の PPTX を変換し、その PDF から互換画像を切り出す。最終 PPTX とは互換画像の中身だけが違う。
    standard, shown = fake_session.renders
    assert standard["input"] == run / "intermediate/final_form.pptx"
    assert (
        standard["input_sha256"]
        == operations["final_form"]["sha256"]
        == operations["render:powerpoint:standard"]["input_sha256"]
    )
    assert operations["math_fallback"]["source_pptx_sha256"] == operations["final_form"]["sha256"]
    assert operations["math_fallback"]["source_pdf_sha256"] == file_hash(run / "renders/powerpoint/standard/deck.pdf")
    assert operations["math_fallback"]["input_form"] == "final_form"
    structure = operations["final_form"]["structure"]
    assert structure["matched"] is True and structure["differences"] == [] and structure["fallback_images"] == 1
    assert shown["input_sha256"] == operations["fallback_clone"]["sha256"]
    # 標準表示の確認の証拠は最終 PPTX の数式の本体（mc:Choice）で決まる。
    parts = pptx_pages(run / "deck.pptx")
    formula = next(page for page in record["slides"] if page["id"] == "formula")
    visual = formula["visuals"]["powerpoint:standard"]
    assert visual["evidence_hash"] == visual_digest(
        parts["formula"], "standard", visual["environment_hash"], render_date=DATE
    )


def test_math_fallback_build_requires_fallback_review_of_math_pages_only(local_project):
    from slidekit.cli import main

    add_math_page(local_project)
    run = built(local_project)
    assert (run / "intermediate/native.pptx").is_file() and (run / "renders/powerpoint/fallback/deck.pdf").is_file()
    presentation = Presentation(run / "deck.pptx")
    alternatives = [slide.element.findall(".//mc:AlternateContent", OOXML) for slide in presentation.slides]
    assert [len(items) for items in alternatives] == [0, 1]
    assert alternatives[1][0].findall("mc:Choice//m:oMath", OOXML)
    blip = alternatives[1][0].find("mc:Fallback//a:blip", OOXML)
    image = presentation.slides[1].part.related_part(blip.get(f"{{{OOXML['r']}}}embed"))
    assert image.blob.startswith(b"\x89PNG")
    # 互換画像は仮の画像（1 色）ではなく、標準表示の PDF から切り出した画像である。
    with Image.open(io.BytesIO(image.blob)) as picture:
        assert len(picture.convert("RGBA").getcolors(1 << 24)) > 2
    record = manifest(run)
    visual = {
        variant: sorted(
            target["slide_id"]
            for target in record["required_reviews"]
            if target["kind"] == "visual" and target["variant"] == variant
        )
        for variant in ("standard", "fallback")
    }
    assert visual == {"standard": ["formula", "intro"], "fallback": ["formula"]}
    operations = record["operations"]
    assert {"final_form", "math_fallback", "fallback_clone"} <= set(operations) and "native_render" not in operations
    assert [crop["slide_id"] for crop in operations["math_fallback"]["crops"]] == ["formula"]
    assert operations["fallback_clone"]["slides"] == [{"number": 2, "id": "formula"}]
    assert operations["fallback_clone"]["kept_slides"] == [{"number": 1, "id": "formula"}]
    # 互換表示の画像と証拠は数式のページだけに作る。画像の名前は元のページ番号で、一覧画像は作らない。
    pages = {page["id"]: page for page in record["slides"]}
    assert {slide_id: sorted(page["visuals"]) for slide_id, page in pages.items()} == {
        "intro": ["powerpoint:standard"],
        "formula": ["powerpoint:fallback", "powerpoint:standard"],
    }
    standard, shown = (pages["formula"]["visuals"][f"powerpoint:{variant}"] for variant in ("standard", "fallback"))
    assert shown["path"] == "renders/powerpoint/fallback/slide-002.png"
    assert shown["pixel_hash"] == pixels(run / shown["path"])
    assert sorted(path.name for path in (run / "renders/powerpoint/fallback").glob("*.png")) == ["slide-002.png"]
    # 標準表示の証拠は数式の本体（mc:Choice）だけで決まり、互換画像の中身は互換表示の証拠にだけ入る。
    parts, native = pptx_pages(run / "deck.pptx"), pptx_pages(run / "intermediate/native.pptx")
    assert pages["formula"]["page_parts"] == parts["formula"].to_json()
    assert parts["formula"].choice == native["formula"].choice and parts["formula"].full != native["formula"].full
    assert standard["evidence_hash"] == visual_digest(
        parts["formula"], "standard", standard["environment_hash"], render_date=DATE
    )
    assert shown["evidence_hash"] == visual_digest(
        parts["formula"], "fallback", shown["environment_hash"], render_date=DATE, number=1
    )
    assert {
        target["evidence_hash"]
        for target in record["required_reviews"]
        if target["kind"] == "visual" and target["slide_id"] == "formula"
    } == {standard["evidence_hash"], shown["evidence_hash"]}
    # 標準表示の確認だけでは、互換表示の確認が残る。
    finish_review(run)
    issues = records.check_build(run, final=True)["issues"]
    assert [
        (issue["code"], issue["target"]["kind"], issue["target"]["slide_id"], issue["target"]["variant"])
        for issue in issues
    ] == [("review_missing", "visual", "formula", "fallback")]
    assert (
        main(
            [
                "review",
                "record",
                "--build",
                str(run),
                "--kind",
                "visual",
                "--slides",
                "formula",
                "--variant",
                "fallback",
                "--reviewer",
                "検査担当",
                "--note",
                "互換表示の数式を確認",
            ]
        )
        == 0
    )
    assert main(["check", "--build", str(run), "--final"]) == 0


def test_final_form_mismatch_is_reported(local_project, monkeypatch):
    add_math_page(local_project)
    original = pipeline.finalize_math

    def altered(final_form, pdf, output, **options):
        # 最終 PPTX のスライドの XML を、最終形の PPTX と違うものにする。
        result = original(final_form, pdf, output, **options)
        parts = read_package(output)
        name = next(name for name in parts if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
        parts[name] = parts[name].replace(b"</p:sld>", b"<!-- changed --></p:sld>")
        write_package(parts, output)
        return {**result, "sha256": file_hash(output)}

    monkeypatch.setattr(pipeline, "finalize_math", altered)
    result = pipeline.build(local_project / "deck.toml")
    assert (result["status"], result["exit_code"]) == ("built", 1)
    (issue,) = [issue for issue in result["issues"] if issue["code"] == "final_form_mismatch"]
    assert issue["severity"] == "error" and issue["observations"][0]["count"] == 1
    structure = manifest(Path(result["build"]))["operations"]["final_form"]["structure"]
    assert structure["matched"] is False and structure["differences"][0]["reason"] == "content"
    assert structure["differences"][0]["part"].startswith("ppt/slides/slide")


def test_fallback_view_renders_only_math_slides(local_project, fake_session):
    rewrite(local_project, pages_source("first", "second", "third", math=("second",)), math=True)
    run = built(local_project)
    with open_pdf(run / "renders/powerpoint/fallback/deck.pdf") as document:
        assert len(document) == 1
    assert sorted(path.name for path in (run / "renders/powerpoint/fallback").glob("*.png")) == ["slide-002.png"]
    assert sorted(path.name for path in (run / "renders/powerpoint/standard").glob("slide-*.png")) == [
        "slide-001.png",
        "slide-002.png",
        "slide-003.png",
    ]
    record = manifest(run)
    assert [
        (target["slide_id"], target["kind"]) for target in record["required_reviews"] if target["variant"] == "fallback"
    ] == [("second", "visual")]
    result_images = pipeline.build(local_project / "deck.toml")["images"]
    assert result_images["powerpoint:fallback"] == ["renders/powerpoint/fallback/slide-002.png"]
    assert result_images["powerpoint:standard"] == ["renders/powerpoint/standard/contact-001.png"]
    # 互換表示の PDF の検査には複製のページだけを渡す（ページ数の不一致にならない）。
    assert "pdf_page_count" not in codes(
        {"issues": json.loads((run / "diagnostics.json").read_text(encoding="utf-8"))["issues"]}
    )


def test_draft_build_skips_fallback(local_project, fake_session):
    rewrite(local_project, pages_source("first", "second", math=("second",)), math=True)
    result = pipeline.build(local_project / "deck.toml", draft=True)
    run = Path(result["build"])
    assert result["exit_code"] == 0, result
    assert len(fake_session.sessions) == 1
    assert [render["label"] for render in fake_session.renders] == ["standard"]
    assert not (run / "intermediate").exists()
    assert "math_fallback_missing" not in codes(result)
    assert (result["mode"], result["exportable"], result["fallback"]) == ("draft", False, False)
    record = manifest(run)
    assert (record["mode"], record["exportable"], record["status"]) == ("draft", False, "built")
    assert {target["variant"] for target in record["required_reviews"] if target["kind"] == "visual"} == {"standard"}


def test_parallel_and_serial_builds_produce_same_images(local_project):
    pages = ("first", "second", "third", "fourth")
    rewrite(local_project, pages_source(*pages, math=pages), math=True)
    serial, parallel = manifest(built(local_project, jobs=1)), manifest(built(local_project, jobs=3))
    assert serial["operations"]["parallel"]["jobs"] == 1 and parallel["operations"]["parallel"]["jobs"] == 3

    def images(record):
        return (
            [{key: visual["pixel_hash"] for key, visual in page["visuals"].items()} for page in record["slides"]],
            [crop["image_sha256"] for crop in record["operations"]["math_fallback"]["crops"]],
        )

    assert images(serial) == images(parallel)
    assert len(images(serial)[1]) == 4
    # 同時プロセス数の誤りは実行ディレクトリを作る前に例外にする。
    with pytest.raises(ValueError, match="--jobs"):
        pipeline.build(local_project / "deck.toml", jobs=0)
    assert sorted(path.name for path in (local_project / "out/runs").iterdir()) == ["0001", "0002"]


def test_checks_follow_conversions_in_order(local_project):
    # PPTX の検査と PDF の検査・画像化は、変換の後に順に行う（別のスレッドで重ねない）。
    rewrite(local_project, pages_source("first", "second", math=("second",)), math=True)
    record = manifest(built(local_project))
    stages = [item["stage"] for item in record["timings"]]
    # 最後の変換の後のセッションの終了（自分で起動した PowerPoint の終了を待つ）も段階に含める。
    order = [
        "final_form",
        "render:powerpoint:standard",
        "math_fallback",
        "fallback_clone",
        "render:powerpoint:fallback",
        "close:powerpoint",
        "pptx_check",
        "pdf:powerpoint:standard",
        "pdf:powerpoint:fallback",
        "collect",
    ]
    assert [stage for stage in stages if stage in order] == order
    assert record["operations"]["parallel"] == {"jobs": 1}
    # 段階は重ならないので、所要時間の合計は全体の時間を超えない。
    assert sum(item["seconds"] for item in record["timings"]) <= record["elapsed_seconds"] + 0.01


def test_manifest_records_timings_and_presentation_seconds(local_project):
    result = pipeline.build(local_project / "deck.toml")
    run = Path(result["build"])
    record = manifest(run)
    stages = [item["stage"] for item in record["timings"]]
    assert {
        "construct",
        "layout",
        "environment",
        "render:powerpoint:standard",
        "close:powerpoint",
        "pdf:powerpoint:standard",
        "collect",
    } <= set(stages)
    assert all(item["seconds"] >= 0 for item in record["timings"]) and record["elapsed_seconds"] > 0
    assert record["presentation_seconds"] == 45 and "duration_seconds" not in record
    assert {item["stage"] for item in result["timings"]} >= set(stages)
    # 結果は CLI がそのまま JSON で出すか、要約にする。
    from slidekit.report import summarize

    checked, examined = pipeline.layout(local_project), pipeline.doctor(local_project)
    for value in (result, checked, examined):
        json.dumps(value, ensure_ascii=False)
    assert summarize("build", result).startswith("生成: 完了（エラー 0 件")
    assert "renders/powerpoint/standard/contact-001.png" in summarize("build", result)
    assert summarize("layout", checked).startswith("配置の検査: 合格")
    assert summarize("doctor", examined).startswith("事前確認: 合格")
    # 予測（layout.json）と、PDF の検査の予測と実際（renders/powerpoint/<表示>/layout.json）を書く。
    predicted = json.loads((run / "layout.json").read_text(encoding="utf-8"))
    assert [element["id"] for element in predicted["slides"][0]["elements"]] == ["body"]
    compared = json.loads((run / "renders/powerpoint/standard/layout.json").read_text(encoding="utf-8"))
    assert (compared["renderer"], compared["variant"]) == ("powerpoint", "standard")
    assert {"layout.json", "renders/powerpoint/standard/layout.json"} <= set(record["artifacts"])


def test_build_shares_one_font_registry(local_project, monkeypatch):
    # 構築・配置の検査・書き出し・PPTX と PDF の検査で 1 つの FontRegistry を使い、最後に 1 回だけ閉じる。
    created, closed = [], []

    class Counted(fonts.FontRegistry):
        def __init__(self, project):
            created.append(project)
            super().__init__(project)

        def close(self):
            closed.append(self)
            super().close()

    monkeypatch.setattr(fonts, "FontRegistry", Counted)
    add_math_page(local_project)
    built(local_project)
    assert len(created) == 1 and len(closed) == 1


# ---------------------------------------------------------------- 部分生成と選んだ章だけの構築


def test_partial_build_renders_only_selected_slides(local_project, fake_session):
    rewrite(local_project, pages_source("first", "second", "third"))
    result = pipeline.build(local_project / "deck.toml", slides="2")
    assert result["exit_code"] == 0, result
    run = Path(result["build"])
    presentation = Presentation(run / "deck.pptx")
    assert len(presentation.slides) == 1
    footer = next(shape for shape in presentation.slides[0].shapes if shape.name == "_theme/footer")
    assert footer.text_frame.text == "2"
    assert sorted(path.name for path in (run / "renders/powerpoint/standard").glob("slide-*.png")) == ["slide-002.png"]
    record = manifest(run)
    assert record["selection"]["slide_ids"] == ["second"] and record["selection"]["total_slides"] == 3
    assert (record["exportable"], result["exportable"], record["mode"]) == (False, False, "full")
    assert [(page["id"], page["number"]) for page in record["slides"]] == [("second", 2)]
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    assert scene["selection"] == record["selection"] and [page["id"] for page in scene["slides"]] == ["second"]
    assert (run / "notes.md").read_text(encoding="utf-8").startswith("## 2. Page 2\n")


def test_partial_build_cannot_be_reviewed_or_exported(local_project):
    rewrite(local_project, pages_source("first", "second", "third"))
    full = built(local_project)
    partial = built(local_project, slides="first")
    with pytest.raises(records.RecordError):
        records.record_review(partial, "visual", "all", "検査担当", "表示を確認")
    warning = make_diagnostic("pdf_font_substitution", "試験", scope="pdf:powerpoint:standard")
    with pytest.raises(records.RecordError):
        records.accept_warning(partial, warning["diagnostic_id"], "検査担当", "試験")
    assert "build_not_exportable" in [issue["code"] for issue in records.check_build(partial, final=True)["issues"]]
    assert records.export_build(partial)["status"] == "failed"
    with pytest.raises(records.RecordError):
        records.carry_reviews(full, partial)


# 章ごとのモジュールを想定した資料。2 つ目の章は書きかけで、作ると例外を出す。選択があれば、選んだページを含む章だけを作る。
CHAPTERS = """from slidekit import Deck, Rect

CHAPTERS = {"first": ("first_intro", "first_detail"), "second": ("second_intro", "second_detail")}
LONG = "Process checked " * 20


def first(deck):
    for page in CHAPTERS["first"]:
        slide = deck.slide(id=page, kind="content", title=page.title(), notes=f"selection={deck.project.selection}")
        text = LONG if page == "first_detail" else "Process checked"
        slide.text(id="body", text=text, area=Rect(60, 110, 400, 60))


def second(deck):
    deck.slide(id="second_intro", kind="content", title="Second")
    raise KeyError("書きかけの章")


def build(context):
    deck = Deck(context)
    for name, chapter in (("first", first), ("second", second)):
        if context.selected(*CHAPTERS[name]):
            chapter(deck)
    return deck
"""


def test_selected_chapters_can_be_checked_while_another_chapter_fails(local_project):
    # 複数人が章ごとに書くとき、ほかの章の書きかけの例外で、自分の章の配置の検査が止まらない。
    rewrite(local_project, CHAPTERS)
    whole = pipeline.layout(local_project)
    assert (whole["status"], whole["exit_code"], codes(whole)) == ("failed", 2, ["build_failed"])
    result = pipeline.layout(local_project, slides="first_intro,first_detail")
    assert (result["status"], result["exit_code"]) == ("failed", 1)
    # 作ったページの中の番号は元のページ番号と違うので、番号を示さず識別子で示す。
    assert result["selection"] == {
        "spec": "first_intro,first_detail",
        "slide_ids": ["first_intro", "first_detail"],
        "numbers": None,
        "total_slides": None,
        "construct": "selected",
    }
    (error,) = [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert (error["code"], error["slide_id"], "slide_number" in error) == ("text_overflow", "first_detail", False)
    assert [(page["id"], page["number"]) for page in result["layout"]["slides"]] == [
        ("first_intro", None),
        ("first_detail", None),
    ]
    from slidekit.report import summarize

    lines = summarize("layout", result).splitlines()
    assert lines[1].startswith("資料: ") and lines[1].endswith(
        "（2 ページ（first_intro、first_detail。選んだ章だけを構築））"
    )
    assert any(line.startswith("  ページ first_detail エラー text_overflow body") for line in lines)
    assert lines[-2].endswith("--slides first_detail")
    # ページ番号を含む指定では、番号を決めるために全ページを作るので、書きかけの章の例外になる。
    numbered = pipeline.layout(local_project, slides="1")
    assert (numbered["exit_code"], codes(numbered)) == (2, ["build_failed"])
    # 識別子の打ち間違いで章を 1 つも作らなかったときは、ページ 0 件の誤りでなく識別子の誤りとして示す。
    # 候補には作らなかった章のページの識別子も入る。
    for run in (
        pipeline.layout(local_project, slides="first_intr"),
        pipeline.build(local_project / "deck.toml", slides="first_intr"),
    ):
        (error,) = [issue for issue in run["issues"] if issue["severity"] == "error"]
        assert (run["exit_code"], error["message"]) == (
            2,
            "ページの識別子がありません: first_intr（候補: first_intro、first_detail）",
        )
    # 選択があっても、資料のコードが選択を使わずにページを作らなければ、これまでどおりページ 0 件の誤りにする。
    rewrite(local_project, "from slidekit import Deck\n\n\ndef build(context):\n    return Deck(context)\n")
    empty = pipeline.layout(local_project, slides="first_intro")
    assert [issue["message"] for issue in empty["issues"]] == ["少なくとも 1 ページを作成してください"]


def test_unknown_slides_are_reported_as_input_error(local_project):
    # --slides のない識別子と範囲外の番号は、利用者の入力の誤り（invalid_selection）として示す。
    # トレースバックを出さず（traceback.txt も書かない）、場所を示さず、次の操作は --slides の指定を直すことにする。
    from slidekit.progress import Progress
    from slidekit.report import summarize

    rewrite(local_project, CHAPTERS)
    message = "ページの識別子がありません: first_intr（候補: first_intro、first_detail）"
    for command in ("layout", "build"):
        stream = io.StringIO()
        if command == "layout":
            result = pipeline.layout(local_project, slides="first_intr", progress=Progress(stream))
        else:
            result = pipeline.build(local_project / "deck.toml", slides="first_intr", progress=Progress(stream))
            assert not (Path(result["build"]) / "traceback.txt").exists()
        assert (result["status"], result["exit_code"]) == ("failed", 2)
        (error,) = result["issues"]
        assert (error["code"], error["severity"], error["message"]) == ("invalid_selection", "error", message)
        assert all("location" not in item and "traceback" not in item for item in error["observations"])
        assert "Traceback" not in stream.getvalue()
        lines = summarize(command, result).splitlines()
        assert f"  エラー invalid_selection {message}" in lines
        assert not any("トレースバック" in line or "slidekit の内部で発生" in line for line in lines)
        assert lines[-2] == f"次の操作: --slides の指定を直して slidekit {command} を再実行する"
    # 範囲外の番号も同じ扱いにする。資料のコードの例外はこれまでどおり場所とトレースバックを示す。
    rewrite(local_project, pages_source("first", "second"))
    result = pipeline.layout(local_project, slides="5")
    assert [(issue["code"], issue["message"]) for issue in result["issues"]] == [
        ("invalid_selection", "ページ番号が範囲外です: 5（1〜2）")
    ]
    rewrite(local_project, CHAPTERS)
    stream = io.StringIO()
    failed = pipeline.layout(local_project, slides="1", progress=Progress(stream))
    (error,) = failed["issues"]
    assert error["code"] == "build_failed" and error["observations"][0]["location"]["in_project"] is True
    assert "Traceback" in stream.getvalue()
    assert summarize("layout", failed).splitlines()[-2] == "次の操作: 原因を直して slidekit layout を再実行する"


def test_selection_is_passed_to_deck_code(local_project):
    project = load_project(local_project)
    assert project.selection is None and project.selected("anything") is True
    project.selection = "intro, id:2024"
    assert project.selected("intro") and project.selected("2024") and project.selected("other", "intro")
    assert not project.selected("other")
    project.selection = "intro,3-4"
    assert project.selected("other") is True
    with pytest.raises(TypeError):
        project.selected()
    # 生成でも同じ指定を渡し、選んだ章だけを作った部分生成は完成物にできない。
    rewrite(local_project, CHAPTERS)
    result = pipeline.build(local_project / "deck.toml", slides="first_intro")
    run = Path(result["build"])
    assert (result["status"], result["exportable"], result["selection"]["construct"]) == ("built", False, "selected")
    record = manifest(run)
    assert record["selection"]["numbers"] is None
    # 作ったページの中の番号は元のページ番号でないので、生成記録・deck.json・ノート・ページ画像では識別子で示す。
    assert [(page["id"], page["number"]) for page in record["slides"]] == [("first_intro", None)]
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    assert [(page["id"], page["number"], page["notes"]) for page in scene["slides"]] == [
        ("first_intro", None, "selection=first_intro")
    ]
    assert (run / "notes.md").read_text(encoding="utf-8").startswith("## first_intro. First_Intro\n")
    assert (
        record["slides"][0]["visuals"]["powerpoint:standard"]["path"]
        == "renders/powerpoint/standard/slide-first_intro.png"
    )
    assert (run / "renders/powerpoint/standard/slide-first_intro.png").is_file()
    # Windows のファイル名で禁止されている文字と % は %XX にする。
    assert (pipeline._image_name(12), pipeline._image_name('a:b%"c')) == ("slide-012.png", "slide-a%3Ab%25%22c.png")
    # 選択を使わない資料（全ページを作る）は、これまでどおり元のページ番号で示す。
    rewrite(local_project, pages_source("first", "second", "third"))
    assert pipeline.layout(local_project, slides="third")["selection"]["construct"] == "all"


def test_waiting_event_names_the_holder():
    # 使用中の記録の Python の PID を優先し、なければ Windows の作業プロセスの PID と判断できる形で示す。
    event = {
        "event": "waiting",
        "waited_seconds": 2.4,
        "holder": {"pid": 5678, "holder": {"project": "/decks/a/deck.toml", "python_pid": 1234}},
    }
    assert (
        pipeline._event_text(event)
        == "変換待ち: 別の処理が PowerPoint を使用中（資料 /decks/a/deck.toml、PID 1234、経過 2 秒）"
    )
    assert pipeline._event_text({"event": "waiting", "holder": {"pid": 5678}}) == (
        "変換待ち: 別の処理が PowerPoint を使用中（Windows の作業プロセスの PID 5678）"
    )
    assert pipeline._event_text({"event": "ready"}) is None


# ---------------------------------------------------------------- 停止と失敗


def test_stop_on_error_skips_preflight_and_conversion(local_project, fake_session):
    rewrite(local_project, pages_source("first", "second", overflow=("second",), math=("first",)), math=True)
    result = pipeline.build(local_project / "deck.toml", stop_on_error=True)
    run = Path(result["build"])
    assert fake_session.sessions == [] and fake_session.renders == []
    assert (result["status"], result["stage"], result["exit_code"]) == ("stopped", "layout", 1)
    assert (run / "deck.pptx").is_file() and (run / "diagnostics.json").is_file() and not (run / "renders").exists()
    assert "text_overflow" in codes(result, "error")
    # 互換画像はまだ作っていないので、その欠落は報告しない。PPTX の検査は行う。
    assert "math_fallback_missing" not in codes(result)
    assert "pptx_check" in [item["stage"] for item in result["timings"]]
    record = manifest(run)
    assert (record["status"], record["stage"]) == ("stopped", "layout") and not (run / ".building").exists()
    assert "build_incomplete" in [issue["code"] for issue in records.check_build(run)["issues"]]


def test_draft_stops_on_error_by_default_and_can_continue(local_project, fake_session):
    rewrite(local_project, pages_source("first", "second", overflow=("second",)))
    stopped = pipeline.build(local_project / "deck.toml", draft=True)
    assert (stopped["status"], stopped["exit_code"]) == ("stopped", 1) and fake_session.renders == []
    continued = pipeline.build(local_project / "deck.toml", draft=True, stop_on_error=False)
    assert (continued["status"], continued["exit_code"], continued["mode"]) == ("built", 1, "draft")
    assert [render["label"] for render in fake_session.renders] == ["standard"]


def test_build_status_matches_manifest_and_counts_quality(local_project):
    rewrite(local_project, pages_source("first", overflow=("first",)))
    result = pipeline.build(local_project / "deck.toml")
    run = Path(result["build"])
    assert (result["status"], manifest(run)["status"], result["exit_code"]) == ("built", "built", 1)
    assert result["counts"]["error"] >= 1 and result["counts"] == {
        severity: sum(issue["severity"] == severity for issue in result["issues"])
        for severity in ("error", "warning", "info")
    }
    (local_project / "slides.py").write_text(
        "def build(context):\n    raise ValueError('資料の入力が不正')\n", encoding="utf-8"
    )
    failed = pipeline.build(local_project / "deck.toml")
    assert (failed["status"], manifest(Path(failed["build"]))["status"], failed["exit_code"]) == ("failed", "failed", 2)


def test_busy_renderer_fails_at_preflight_with_layout_diagnostics_kept(local_project, fake_session):
    rewrite(local_project, pages_source("first", overflow=("first",)))

    def busy(on_event):
        # 使用中の記録（active.json）は、Windows の作業プロセスの PID と、その処理を始めた側の情報を持つ。
        on_event(
            {
                "event": "waiting",
                "waited_seconds": 5.0,
                "holder": {"pid": 5678, "holder": {"run": "/other/out/runs/0003", "python_pid": 1234}},
            }
        )
        raise RendererError(
            "SLIDEKIT_BUSY: 同じ Windows 利用者による変換が実行中です",
            code="renderer_busy",
            holder={"run": "/other/out/runs/0003"},
            waited_seconds=5.0,
        )

    fake_session.on_open = busy
    from slidekit.progress import Progress

    stream = io.StringIO()
    result = pipeline.build(local_project / "deck.toml", progress=Progress(stream), wait_seconds=3)
    run = Path(result["build"])
    assert fake_session.sessions[0]["wait_seconds"] == 3
    assert (result["status"], result["stage"], result["exit_code"]) == ("failed", "doctor", 2)
    assert {"text_overflow", "renderer_busy"} <= set(codes(result))
    assert (manifest(run)["status"], manifest(run)["stage"]) == ("failed", "doctor")
    assert (
        "変換待ち: 別の処理が PowerPoint を使用中（実行 /other/out/runs/0003、PID 1234、経過 5 秒）"
        in stream.getvalue()
    )
    assert (run / "deck.pptx").is_file() and (run / "traceback.txt").is_file()


def test_renderer_failure_keeps_manifest_and_diagnostics(local_project, fake_session):
    def fail(label, source, output):
        raise RendererError("表示ソフトの制限時間を超えた", code="renderer_timeout")

    fake_session.on_render = fail
    result = pipeline.build(local_project / "deck.toml")
    assert result["exit_code"] == 2
    run = Path(result["build"])
    assert (run / "deck.pptx").is_file()
    assert not (run / ".building").exists()
    record = manifest(run)
    assert (record["status"], record["stage"]) == ("failed", "render:powerpoint:standard")
    assert record["schema_version"] == records.SCHEMA_VERSION
    assert "renderer_timeout" in codes(result)
    # 開いたセッションは、失敗しても閉じる。
    assert fake_session.closed == 1
    assert records.check_build(run)["status"] == "failed"
    assert records.export_build(run)["status"] == "failed"
    assert not (local_project / "presentation.pptx").exists()


def test_interrupted_build_is_recorded_as_failed(local_project, fake_session):
    def interrupt(label, source, output):
        raise KeyboardInterrupt

    fake_session.on_render = interrupt
    with pytest.raises(KeyboardInterrupt) as raised:
        pipeline.build(local_project / "deck.toml")
    (run,) = (local_project / "out/runs").iterdir()
    record = manifest(run)
    assert (record["status"], record["stage"]) == ("failed", "render:powerpoint:standard")
    # コマンドが実行ディレクトリを示すことができるように、その時点の戻り値を例外に付けて送り直す。
    result = raised.value.result
    assert (result["build"], result["status"], result["stage"], result["exit_code"]) == (
        str(run),
        "failed",
        "render:powerpoint:standard",
        2,
    )
    assert "build_interrupted" in codes(result, "error") and "carry" not in result
    issues = json.loads((run / "diagnostics.json").read_text(encoding="utf-8"))["issues"]
    assert [issue["code"] for issue in issues if issue["severity"] == "error"] == ["build_interrupted"]
    assert not (run / ".building").exists() and fake_session.closed == 1


def test_pdf_check_failure_names_its_stage(local_project, monkeypatch):
    # 変換の後に行う PDF の検査の例外は、その段階の失敗として記録する。
    add_math_page(local_project)
    check = pipeline.validate_pdf

    def broken(pdf, scene, project, renderer, variant, **options):
        if variant == "standard":
            raise RuntimeError("PDF を読めない")
        return check(pdf, scene, project, renderer, variant, **options)

    monkeypatch.setattr(pipeline, "validate_pdf", broken)
    result = pipeline.build(local_project / "deck.toml")
    assert (result["status"], result["stage"], result["exit_code"]) == ("failed", "pdf:powerpoint:standard", 2)
    (issue,) = [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert (issue["code"], issue["scope"]) == ("build_failed", "pdf:powerpoint:standard")
    assert not (Path(result["build"]) / "intermediate/fallback.pptx").exists()


def test_collect_failure_names_its_stage(local_project, monkeypatch):
    # 検査の後の集約（deck.pdf のリンクを含む）の例外を、終わった検査の段階の失敗にしない。
    add_math_page(local_project)

    def broken(source, target):
        raise RuntimeError("deck.pdf を作れない")

    monkeypatch.setattr(pipeline, "_link", broken)
    result = pipeline.build(local_project / "deck.toml")
    assert (result["status"], result["stage"], result["exit_code"]) == ("failed", "collect", 2)
    (issue,) = [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert (issue["code"], issue["scope"]) == ("build_failed", "collect")
    assert manifest(Path(result["build"]))["stage"] == "collect"


def test_environment_failure_names_its_stage(local_project, monkeypatch):
    # 実行環境の記録の例外を、終わった書き出しや配置の検査の段階の失敗にしない（build・doctor とも）。
    def broken(project, registry):
        raise RuntimeError("書体のファイルを読めない")

    monkeypatch.setattr(pipeline, "_environment", broken)
    result = pipeline.build(local_project / "deck.toml")
    assert (result["status"], result["stage"], result["exit_code"]) == ("failed", "environment", 2)
    (issue,) = [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert (issue["code"], issue["scope"]) == ("build_failed", "environment")
    examined = pipeline.doctor(local_project)
    (issue,) = [issue for issue in examined["issues"] if issue["severity"] == "error"]
    assert examined["exit_code"] == 2 and (issue["code"], issue["scope"]) == ("build_failed", "environment")


def test_construction_failure_still_has_checkable_run(local_project):
    (local_project / "slides.py").write_text(
        "def build(context):\n    raise ValueError('資料の入力が不正')\n", encoding="utf-8"
    )
    result = pipeline.build(local_project / "deck.toml")
    assert result["exit_code"] == 2
    run = Path(result["build"])
    assert manifest(run)["stage"] == "construct"
    assert records.check_build(run, final=True)["status"] == "failed"
    assert not (run / ".building").exists()


def test_construct_error_reports_file_line_and_traceback(local_project):
    lines = ["# 説明の行"] * 54 + [
        "def build(context):",
        '    labels = {"a": 1}',
        '    label = labels["b"]',
        "    return label",
    ]
    (local_project / "slides.py").write_text("\n".join(lines) + "\n", encoding="utf-8")
    result = pipeline.build(local_project / "deck.toml")
    run = Path(result["build"])
    (issue,) = [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert (issue["code"], issue["scope"]) == ("build_failed", "construct")
    observed = issue["observations"][0]
    assert observed["exception"] == "KeyError" and observed["traceback"] == "traceback.txt"
    assert (observed["location"]["line"], observed["location"]["relative_path"]) == (57, "slides.py")
    assert observed["location"]["code"] == 'label = labels["b"]' and observed["location"]["in_project"] is True
    assert "KeyError" in (run / "traceback.txt").read_text(encoding="utf-8")
    assert "traceback.txt" in manifest(run)["artifacts"]
    # 配置だけの検査は実行ディレクトリがないので、トレースバックを書かずに失敗の診断を返す。
    checked = pipeline.layout(local_project)
    assert (checked["status"], checked["exit_code"], codes(checked)) == ("failed", 2, ["build_failed"])


def test_input_error_ids_are_copied_to_diagnostic(local_project):
    (local_project / "slides.py").write_text(
        """from slidekit.errors import InputError

def build(context):
    raise InputError("ページ intro の要素 body: 値が不正です", code="invalid_style", slide_id="intro", element_id="body")
""",
        encoding="utf-8",
    )
    result = pipeline.build(local_project / "deck.toml")
    (issue,) = [issue for issue in result["issues"] if issue["severity"] == "error"]
    assert (issue["code"], issue["slide_id"], issue["element_id"]) == ("invalid_style", "intro", "body")
    assert result["exit_code"] == 2


def test_deferred_construct_issues_are_merged(local_project, monkeypatch):
    # 作成 API が続行できる誤りを集めるのは、生成関数が呼ばれる時点で collect_errors が真のときだけである。
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace(
            'notes="確認内容を説明します。"', "notes=f\"collect_errors={getattr(context, 'collect_errors', None)}\""
        ),
        encoding="utf-8",
    )
    deferred = make_diagnostic(
        "invalid_style",
        "ページ intro の要素 body: 書式の値が不正です（size_pt: 'large'）",
        scope="input",
        slide_id="intro",
        element_id="body",
        observations={"size_pt": "large"},
    )
    original = Deck.to_dict

    def to_dict(self):
        scene = original(self)
        scene["issues"] = [deferred]
        return scene

    monkeypatch.setattr(Deck, "to_dict", to_dict)
    result = pipeline.build(local_project / "deck.toml")
    run = Path(result["build"])
    assert result["exit_code"] == 1 and result["stage"] == "complete"
    assert deferred in result["issues"]
    assert deferred in json.loads((run / "diagnostics.json").read_text(encoding="utf-8"))["issues"]
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    assert scene["issues"] == [deferred]
    assert scene["slides"][0]["notes"] == "collect_errors=True"
    assert manifest(run)["status"] == "built"
    # 配置だけの事前確認も、集めた誤りを配置の診断と合わせて示す。
    checked = pipeline.doctor(local_project)
    assert deferred in checked["issues"] and checked["exit_code"] == 1


def test_layout_diagnostics_exist_before_conversion(local_project, fake_session):
    rewrite(local_project, pages_source("first", overflow=("first",)))
    seen = {}

    def inspect():
        run = Path(fake_session.sessions[0]["holder"]["run"])
        data = json.loads((run / "diagnostics.json").read_text(encoding="utf-8"))
        seen.update(
            provisional=data.get("provisional"),
            codes=[issue["code"] for issue in data["issues"]],
            files={name for name in ("deck.json", "notes.md", "layout.json") if (run / name).is_file()},
            stage=json.loads((run / ".building").read_text(encoding="utf-8"))["stage"],
        )

    fake_session.on_doctor = inspect
    result = pipeline.build(local_project / "deck.toml")
    assert seen["provisional"] is True and "text_overflow" in seen["codes"]
    assert seen["files"] == {"deck.json", "notes.md", "layout.json"} and seen["stage"] == "doctor"
    final = json.loads((Path(result["build"]) / "diagnostics.json").read_text(encoding="utf-8"))
    assert "provisional" not in final


def test_missing_bibtexparser_is_explained(local_project, monkeypatch):
    (local_project / "references.bib").write_text("@book{knuth1997, title = {Book}}\n", encoding="utf-8")
    settings = local_project / "deck.toml"
    settings.write_text(
        settings.read_text(encoding="utf-8").replace(
            'read_roots = ["."]', 'read_roots = ["."]\nbibliography = ["references.bib"]'
        ),
        encoding="utf-8",
    )
    monkeypatch.setitem(sys.modules, "bibtexparser", None)
    result = pipeline.build(local_project / "deck.toml")
    assert result["exit_code"] == 2
    assert any("bibliography" in issue["message"] for issue in result["issues"] if issue["severity"] == "error")


# ---------------------------------------------------------------- 互換画像の切り出しと許容の宣言


CLIPPED = ['    slide.text(id="equation", text=[Math("888")], area=Rect(60, 100, 100, 130), size_pt=100)']


def test_clipped_crop_reports_warning(local_project):
    # 偽の変換は数式の文字を枠の左上から描くので、大きな数字の字形が互換画像の範囲（枠から 16 pt）を超える。
    rewrite(local_project, pages_source("first", extra={"first": CLIPPED}), math=True)
    result = pipeline.build(local_project / "deck.toml")
    assert result["status"] == "built"
    (warning,) = [issue for issue in result["issues"] if issue["code"] == "fallback_crop_clipped"]
    assert (warning["severity"], warning["scope"], warning["slide_id"], warning["element_id"]) == (
        "warning",
        "pptx",
        "first",
        "equation",
    )
    (crop,) = manifest(Path(result["build"]))["operations"]["math_fallback"]["crops"]
    assert crop["clipped"] is True and warning["observations"][0]["crop_pt"] == crop["crop_pt"]
    assert "declared" not in warning
    # 資料コードで許容を宣言すると、生成の後に出た警告にも宣言が当たる（最後の診断の結合の後に当てる）。
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace(
            CLIPPED[0],
            CLIPPED[0]
            + '\n    slide.allow_warning("fallback_crop_clipped", element="equation", reason="互換表示で欠けを確かめた")',
        ),
        encoding="utf-8",
    )
    declared = pipeline.build(local_project / "deck.toml")
    (warning,) = [issue for issue in declared["issues"] if issue["code"] == "fallback_crop_clipped"]
    assert warning["declared"]["reason"] == "互換表示で欠けを確かめた"
    assert "allow_warning_unused" not in codes(declared)


def test_unused_warning_declaration_is_reported(local_project):
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace(
            "    return deck", '    slide.allow_warning("pdf_font_substitution", reason="試験の宣言")\n    return deck'
        ),
        encoding="utf-8",
    )
    result = pipeline.build(local_project / "deck.toml")
    assert ("allow_warning_unused", "info") in [(issue["code"], issue["severity"]) for issue in result["issues"]]
    assert result["exit_code"] == 0


# ---------------------------------------------------------------- 実行ディレクトリ


def test_primary_pdf_is_stored_once(local_project):
    run = built(local_project)
    primary, standard = run / "deck.pdf", run / "renders/powerpoint/standard/deck.pdf"
    try:
        assert os.stat(primary).st_ino == os.stat(standard).st_ino
    except AssertionError:
        # ハードリンクを作成できないファイルシステムでは複製する。
        assert primary.read_bytes() == standard.read_bytes()
    record = manifest(run)
    assert (
        record["artifacts"]["deck.pdf"]
        == record["artifacts"]["renders/powerpoint/standard/deck.pdf"]
        == file_hash(primary)
    )


def test_fallback_clone_is_removed_unless_kept(local_project):
    add_math_page(local_project)
    run = built(local_project)
    record = manifest(run)
    assert not (run / "intermediate/fallback.pptx").exists() and not (run / "intermediate/final_form.pptx").exists()
    assert record["operations"]["fallback_clone"]["sha256"] and record["operations"]["fallback_clone"]["kept"] is False
    assert not {"intermediate/fallback.pptx", "intermediate/final_form.pptx"} & set(record["artifacts"])
    kept = built(local_project, keep_intermediate=True)
    record = manifest(kept)
    assert file_hash(kept / "intermediate/fallback.pptx") == record["operations"]["fallback_clone"]["sha256"]
    assert file_hash(kept / "intermediate/final_form.pptx") == record["operations"]["final_form"]["sha256"]
    assert record["artifacts"]["intermediate/fallback.pptx"] == record["operations"]["fallback_clone"]["sha256"]
    # 成果物の記録は実際のファイルの内容と一致する（途中で求めた値を使い回しても）。
    assert all(file_hash(kept / name) == digest for name, digest in record["artifacts"].items())


def test_sequential_builds_allocate_new_immutable_directories(local_project):
    first = built(local_project)
    snapshot = {str(path.relative_to(first)): path.read_bytes() for path in first.rglob("*") if path.is_file()}
    second = built(local_project)
    assert first.name == "0001" and second.name == "0002"
    assert snapshot == {str(path.relative_to(first)): path.read_bytes() for path in first.rglob("*") if path.is_file()}
    # 実際の生成が書いた記録でも、0.1 の呼び方の records.clean_runs（runs.clean_runs を呼ぶ互換の関数）で古い実行が整理の候補になる。
    preview = records.clean_runs(load_project(local_project), "0002")
    assert preview["candidates"] == [str(first)]


# ---------------------------------------------------------------- ノート


def test_notes_heading_joins_title_lines():
    assert pipeline._heading_text("資料の\n作り方") == "資料の作り方"
    assert pipeline._heading_text("Process\nchecked") == "Process checked"
    assert pipeline._heading_text("事例\nA") == "事例 A"
    assert pipeline._heading_text("  前後の空白 \n\n  を除く ") == "前後の空白を除く"
    assert (
        pipeline._heading_text([{"kind": "text", "text": "面積 "}, {"kind": "math", "text": "\\pi r^2"}])
        == "面積 $\\pi r^2$"
    )
    scene = {
        "slides": [
            {"id": "a", "number": 3, "title": "資料の\n作り方", "notes": "説明します。"},
            {"id": "b", "number": 5, "title": "", "notes": ""},
        ]
    }
    assert pipeline.notes_markdown(scene) == "## 3. 資料の作り方\n\n説明します。\n\n## 5.\n\n\n"


# ---------------------------------------------------------------- 記録・確認・引き継ぎ・反映


def test_fallback_evidence_uses_position_in_review_copy(local_project):
    """互換表示の確認用複製は数式のページだけにし、互換表示の証拠の番号は複製の中の位置にする。

    雛形のマスターにページ番号のフィールドがあると、確認用複製に表示される番号は複製の中の位置になる。
    資料の中の位置を証拠に使うと、前のページの数式を加えたり除いたりして番号が変わっても、確認が引き継がれてしまう。
    """
    template = Presentation()
    template.slide_master.element.find("p:cSld/p:spTree", OOXML).append(etree.fromstring(NUMBER_FIELD))
    template.save(local_project / "numbered.pptx")
    add_math_page(local_project)
    settings = local_project / "deck.toml"
    settings.write_text(
        settings.read_text(encoding="utf-8")
        + '[template]\npath = "numbered.pptx"\nlayout_index = 6\nshow_master_shapes = true\n',
        encoding="utf-8",
    )
    run = built(local_project)
    record = manifest(run)
    assert record["operations"]["fallback_clone"]["kept_slides"] == [{"number": 1, "id": "formula"}]
    formula = next(page for page in record["slides"] if page["id"] == "formula")
    shown = formula["visuals"]["powerpoint:fallback"]
    # PDF の 1 ページ目だが、画像の名前は元のページ番号にする。
    assert shown["path"] == "renders/powerpoint/fallback/slide-002.png"
    parts = pptx_pages(run / "deck.pptx")["formula"]
    assert parts.fields == ("slidenum",) and parts.number == 2
    in_copy = visual_digest(parts, "fallback", shown["environment_hash"], render_date=DATE, number=1)
    assert shown["evidence_hash"] == in_copy
    assert in_copy != visual_digest(parts, "fallback", shown["environment_hash"], render_date=DATE, number=2)
    # 標準表示は資料の中の位置のまま。
    standard = formula["visuals"]["powerpoint:standard"]
    assert standard["evidence_hash"] == visual_digest(
        parts, "standard", standard["environment_hash"], render_date=DATE, number=2
    )


# 雛形のマスターに置く、ページ番号のフィールドを持つ通常の図形（表示される）。
NUMBER_FIELD = (
    f'<p:sp xmlns:p="{OOXML["p"]}" xmlns:a="{OOXML["a"]}"><p:nvSpPr><p:cNvPr id="90" name="Page number"/>'
    '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="8382000" y="6400800"/>'
    '<a:ext cx="609600" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
    "<a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>"
)


def test_real_build_review_and_export(local_project, dejavu):
    root = local_project
    (root / "brief.md").write_text("結合検査の目的を確かめる。\n", encoding="utf-8")
    source_before = (root / "slides.py").read_bytes()
    run = built(root)
    assert not (run / ".building").exists()
    assert not (root / "presentation.pptx").exists()
    record = manifest(run)
    # 形式 2 の生成記録。全ページの通常の生成なので完成物にできる。
    assert (record["schema_version"], record["evidence_version"]) == (records.SCHEMA_VERSION, EVIDENCE_VERSION)
    assert [record[key] for key in ("status", "stage", "mode", "selection", "exportable")] == [
        "built",
        "complete",
        "full",
        None,
        True,
    ]
    assert record["presentation_seconds"] == 45 and "duration_seconds" not in record
    assert "renders/powerpoint/standard/slide-001.png" in record["artifacts"]
    assert not {"manifest.json", "review.json", "check.json", "final_check.json", ".building"} & set(
        record["artifacts"]
    )
    assert any(name.endswith("/pipeline.py") for name in record["input_files"])
    assert record["environment"]["fonts"]["fonts"]["DejaVu Sans"]["sha256"]
    assert "matplotlib" not in record["environment"]["packages"]
    for name in ("diagnostics.json", "review.json"):
        assert json.loads((run / name).read_text(encoding="utf-8"))["schema_version"] == records.SCHEMA_VERSION
    # 証拠は成果物と構築結果から同じ関数で求め直した値と一致する。
    project = load_project(root)
    project.config["_bibliography_entries"] = {}
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    parts = pptx_pages(run / "deck.pptx")
    (page,) = record["slides"]
    assert (page["id"], page["number"]) == ("intro", 1)
    assert page["content_hash"] == content_digest(scene["slides"][0], project)
    # 依頼内容（brief.md）は資料全体の証拠に入る（ページの内容の証拠には入らない）。
    assert record["input_files"][str(root / "brief.md")] == file_hash(root / "brief.md")
    assert record["deck_content_hash"] == deck_digest(scene, project)
    project.brief_path = None
    assert record["deck_content_hash"] != deck_digest(scene, project)
    assert page["content_hash"] == content_digest(scene["slides"][0], project)
    assert page["page_parts"] == parts["intro"].to_json()
    visual = page["visuals"]["powerpoint:standard"]
    assert visual["path"] == "renders/powerpoint/standard/slide-001.png" and visual["pixel_hash"] == pixels(
        run / visual["path"]
    )
    # 表示環境のフォントは実行環境で宣言した書体の内容（計測で使った書体の一覧ではない）。
    declared = {
        "DejaVu Sans": {
            "sha256": file_hash(dejavu.path),
            "index": 0,
            "bold_sha256": file_hash(dejavu.bold_path),
            "bold_index": 0,
        }
    }
    environment = display_environment(
        record["environment"]["renderers"]["powerpoint"],
        record["operations"]["render:powerpoint:standard"],
        declared,
        record["environment"]["packages"]["pypdfium2"],
    )
    assert visual["environment_hash"] == environment
    assert visual["evidence_hash"] == visual_digest(parts["intro"], "standard", environment, render_date=DATE)
    contact = pixels(run / "renders/powerpoint/standard/contact-001.png")
    assert {(target["kind"], target["slide_id"]): target["evidence_hash"] for target in record["required_reviews"]} == {
        ("content", "intro"): page["content_hash"],
        ("visual", "intro"): visual["evidence_hash"],
        ("overview", None): overview_digest([contact], environment, record["deck_content_hash"]),
    }
    pptx = Presentation(run / "deck.pptx")
    assert pptx.core_properties.title == "結合検査"
    assert pptx.core_properties.author == ""
    assert pptx.core_properties.last_modified_by == ""
    assert pptx.slides[0].notes_slide.notes_text_frame.text == "確認内容を説明します。"
    assert records.check_build(run)["status"] == "passed"
    assert records.check_build(run, final=True)["status"] == "failed"
    assert records.export_build(run)["status"] == "failed"
    finish_review(run)
    assert records.check_build(run, final=True)["status"] == "passed"
    result = records.export_build(run)
    assert result["status"] == "exported"
    assert (run / "final_check.json").is_file()
    assert (root / "presentation.pptx").read_bytes() == (run / "deck.pptx").read_bytes()
    assert (root / "presentation.pdf").read_bytes() == (run / "deck.pdf").read_bytes()
    assert (root / "slides.py").read_bytes() == source_before


def review_kinds(run):
    """実行の確認記録の種類（記録の順）。"""
    return [record["kind"] for record in json.loads((run / "review.json").read_text(encoding="utf-8"))["records"]]


def test_same_second_same_length_edit_uses_current_source(local_project):
    source = local_project / "slides.py"
    first_source = source.read_text(encoding="utf-8").replace("確認内容を説明します。", "最初の内容を説明します。")
    source.write_text(first_source, encoding="utf-8")
    fixed_time = 1700000000
    os.utime(source, (fixed_time, fixed_time))
    first = built(local_project)
    source.write_text(first_source.replace("最初の内容", "次回の内容"), encoding="utf-8")
    os.utime(source, (fixed_time, fixed_time))
    second = built(local_project)
    assert "次回の内容" in (second / "notes.md").read_text(encoding="utf-8")
    assert manifest(first)["slides"][0]["content_hash"] != manifest(second)["slides"][0]["content_hash"]


def test_note_change_carries_visual_only(local_project):
    # ノートだけを変えて再生成すると、表示の確認だけが自動で引き継がれ、内容と一覧画像の確認が残る。
    first = built(local_project)
    finish_review(first)
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace("確認内容を説明します。", "別の内容を説明します。"), encoding="utf-8"
    )
    result = pipeline.build(local_project / "deck.toml")
    second = Path(result["build"])
    assert (result["exit_code"], result["carry"]["status"], result["carry"]["count"]) == (0, "carried", 1)
    assert result["carry"]["remaining"] == {"visual": 0, "overview": 1, "content": 1, "warnings": 0}
    assert review_kinds(second) == ["visual"]
    status = records.review_status(second)
    assert [(group["kind"], group["satisfied"], group["carried"]) for group in status["reviews"]] == [
        ("visual", 1, 1),
        ("overview", 0, 0),
        ("content", 0, 0),
    ]


def test_carry_from_none_keeps_new_run_empty(local_project):
    first = built(local_project)
    finish_review(first)
    result = pipeline.build(local_project / "deck.toml", carry_from="none")
    assert "carry" not in result and review_kinds(Path(result["build"])) == []
    # 実行を指定すると、その実行からだけ引き継ぐ（番号でも指定できる）。
    chosen = pipeline.build(local_project / "deck.toml", carry_from="1")
    assert (chosen["carry"]["status"], chosen["carry"]["mode"], chosen["carry"]["count"]) == ("carried", "manual", 2)
    assert chosen["carry"]["sources"] == {str(first): 2}
    assert sorted(review_kinds(Path(chosen["build"]))) == ["content", "visual"]
    # latest は生成の前の最新の実行を指す（生成した実行自身ではない）。
    latest = pipeline.build(local_project / "deck.toml", carry_from="latest")
    assert (latest["carry"]["status"], latest["carry"]["count"]) == ("carried", 2)
    assert latest["carry"]["sources"] == {chosen["build"]: 2}


def test_draft_build_is_not_carried_into(local_project):
    first = built(local_project)
    finish_review(first)
    for options in ({"draft": True}, {"slides": "intro"}):
        result = pipeline.build(local_project / "deck.toml", **options)
        assert result["status"] == "built" and result["exportable"] is False
        assert "carry" not in result and review_kinds(Path(result["build"])) == []


def test_carry_failure_does_not_fail_build(local_project, monkeypatch):
    # 書式の誤りは実行ディレクトリを作る前に知らせる。
    with pytest.raises(ValueError, match="carry_from"):
        pipeline.build(local_project / "deck.toml", carry_from="atuo")
    assert not (local_project / "out/runs").exists() or not any((local_project / "out/runs").iterdir())
    # 引き継ぎの失敗（完成物にできる実行がない、指定した実行がない、ロックを取得できない）は、生成の失敗にせず、終了コードも変えない。
    nothing = pipeline.build(local_project / "deck.toml", carry_from="latest")
    assert (nothing["status"], nothing["exit_code"], nothing["carry"]["status"]) == ("built", 0, "failed")
    assert "完成物にできる実行がない" in nothing["carry"]["message"]
    missing = pipeline.build(local_project / "deck.toml", carry_from="9")
    assert (missing["status"], missing["exit_code"], missing["carry"]["status"]) == ("built", 0, "failed")
    assert "9" in missing["carry"]["message"] or "0009" in missing["carry"]["message"]

    def locked(target, **options):
        return {
            "command": "review carry",
            "status": "skipped",
            "exit_code": 0,
            "build": str(target),
            "message": "別の処理が実行中のため引き継がなかった",
        }

    monkeypatch.setattr(records, "carry_reviews_auto", locked)
    skipped = pipeline.build(local_project / "deck.toml")
    assert skipped["exit_code"] == 0
    assert skipped["carry"] == {
        "status": "skipped",
        "mode": "auto",
        "message": "別の処理が実行中のため引き継がなかった",
    }


def test_review_on_stale_run_carries_after_rebuild(local_project):
    # 入力を変えた後でも古い実行に確認を記録でき、変えていないページの確認は再生成で引き継がれる。
    rewrite(local_project, pages_source("first", "second"))
    stale = built(local_project)
    source = local_project / "slides.py"
    source.write_text(
        pages_source(
            "first",
            "second",
            extra={"second": ['    slide.text(id="added", text="Added line", area=Rect(60, 300, 400, 60))']},
        ),
        encoding="utf-8",
    )
    recorded = records.record_review(stale, "visual,content", "all", "検査担当", "表示と本文を確認")
    assert [notice["code"] for notice in recorded["notices"]] == ["input_changed"]
    result = pipeline.build(local_project / "deck.toml")
    copies = json.loads((Path(result["build"]) / "review.json").read_text(encoding="utf-8"))["records"]
    assert sorted((record["kind"], record["slide_id"]) for record in copies) == [
        ("content", "first"),
        ("visual", "first"),
    ]
    assert result["carry"]["remaining"] == {"visual": 1, "overview": 1, "content": 1, "warnings": 0}


def test_page_insert_carries_unchanged_content_and_same_digit_visuals(local_project):
    # ページ番号のある 3 ページを確認し、先頭に 1 ページ加えて再生成すると、既存の 3 ページの内容と表示の確認が引き継がれる
    # （番号の桁数が同じなので、番号の領域を除いて表示を比べる）。未確認は新しいページと一覧画像だけになる。
    # 数式のページは標準表示と互換表示の両方の確認を引き継ぐが、番号だけ変わったページとしては 1 回だけ挙げる。
    titles = {page: page.capitalize() for page in ("zero", "first", "second", "third")}
    rewrite(local_project, pages_source("first", "second", "third", titles=titles, math=("second",)), math=True)
    first = built(local_project)
    finish_review(first)
    records.record_review(first, "visual", "all", "検査担当", "互換表示を確認", variant="fallback")
    rewrite(local_project, pages_source("zero", "first", "second", "third", titles=titles, math=("second",)))
    result = pipeline.build(local_project / "deck.toml")
    second = Path(result["build"])
    assert result["carry"]["count"] == 7 and result["carry"]["renumbered"] == ["first", "second", "third"]
    copies = json.loads((second / "review.json").read_text(encoding="utf-8"))["records"]
    assert sorted(
        (record["slide_id"], record["variant"], bool(record["render_check"]["masked"]))
        for record in copies
        if record["kind"] == "visual"
    ) == [
        ("first", "standard", True),
        ("second", "fallback", True),
        ("second", "standard", True),
        ("third", "standard", True),
    ]
    from slidekit.report import summarize

    # 引き継ぎ元と、ページ番号だけ変わったページを比べた実行を示す。
    first_run = Path(result["carry"]["compared_run"] or "0001").name
    assert result["carry"]["renumbered_from"] == {first_run: 3}
    assert any(
        line.startswith(
            f"確認の引き継ぎ: 7 件（引き継ぎ元 {first_run} 7 件、ページ番号だけ変わったページ 3 件（実行 {first_run} と比較）、"
        )
        for line in summarize("build", result).splitlines()
    )
    status = records.review_status(second)
    pending = {
        (group["kind"], tuple(item["slide_id"] for item in group["pending"]))
        for group in status["reviews"]
        if group["pending"]
    }
    assert pending == {("visual", ("zero",)), ("content", ("zero",)), ("overview", (None,))}


def test_brief_change_requires_only_overview(local_project):
    # 依頼内容（brief.md）は資料全体の証拠にだけ入るので、変えて再生成すると一覧画像の確認だけが残る。
    (local_project / "brief.md").write_text("結合検査の目的を確かめる。\n", encoding="utf-8")
    first = built(local_project)
    finish_review(first)
    (local_project / "brief.md").write_text("結合検査の目的と範囲を確かめる。\n", encoding="utf-8")
    result = pipeline.build(local_project / "deck.toml")
    assert result["carry"]["remaining"] == {"visual": 0, "overview": 1, "content": 0, "warnings": 0}
    assert sorted(review_kinds(Path(result["build"]))) == ["content", "visual"]


def test_converted_asset_paths_do_not_prevent_unchanged_content_carry(local_project):
    asset = local_project / "diagram.pdf"
    with Document() as document:
        page = document.new_page(width=100, height=100)
        page.draw_rect((10, 10, 90, 90), color=(1, 0, 0), fill=(1, 0, 0))
        document.save(asset)
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace(
            "    return deck",
            """    slide.image(id="figure", path="diagram.pdf", page=1, area=Rect(60, 240, 100, 100))
    return deck""",
        ),
        encoding="utf-8",
    )
    first = built(local_project)
    finish_review(first)
    result = pipeline.build(local_project / "deck.toml")
    second = Path(result["build"])
    old, new = manifest(first), manifest(second)
    # 変換した画像の置き場所は実行ごとに違うが、内容と表示の証拠は同じなので、2 回目の build だけで表示と内容の確認が
    # 自動で引き継がれる。
    assert old["slides"][0]["content_hash"] == new["slides"][0]["content_hash"]
    assert old["slides"][0]["visuals"] == new["slides"][0]["visuals"]
    carry = result["carry"]
    assert (result["exit_code"], carry["status"], carry["mode"], carry["count"], carry["sources"]) == (
        0,
        "carried",
        "auto",
        2,
        {str(first): 2},
    )
    assert carry["remaining"] == {"visual": 0, "overview": 1, "content": 0, "warnings": 0}
    assert sorted(review_kinds(second)) == ["content", "visual"]
    # 写しは人が記録した元の記録を指す。
    copies = json.loads((second / "review.json").read_text(encoding="utf-8"))["records"]
    assert {record["source"]["run"] for record in copies} == {str(first)}
    # 一覧画像の確認は引き継がないので、最終検査の残りは一覧画像の確認だけになり、それを記録すると反映できる。
    result = records.check_build(second, final=True)
    assert [(issue["code"], issue["target"]["kind"]) for issue in result["issues"]] == [("review_missing", "overview")]
    records.record_review(second, "overview", "all", "検査担当", "新しい一覧を確認")
    assert records.export_build(second)["status"] == "exported"


def test_asset_replacement_invalidates_content_and_visual_confirmation(local_project):
    image = local_project / "diagram.png"
    Image.new("RGB", (40, 40), "red").save(image)
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace(
            "    return deck",
            """    slide.image(id="figure", path="diagram.png", area=Rect(60, 240, 100, 100))
    return deck""",
        ),
        encoding="utf-8",
    )
    first = built(local_project)
    finish_review(first)
    Image.new("RGB", (40, 40), "blue").save(image)
    result = pipeline.build(local_project / "deck.toml")
    second = Path(result["build"])
    old, new = manifest(first), manifest(second)
    assert old["slides"][0]["content_hash"] != new["slides"][0]["content_hash"]
    assert old["slides"][0]["visuals"] != new["slides"][0]["visuals"]
    assert result["carry"]["count"] == 0 and review_kinds(second) == []
    assert records.carry_reviews(first, second)["records"] == []


# ---------------------------------------------------------------- 資料コードの読み込み


def constructed(root):
    """資料を読み込んで構築し、最初のページのノートと入力の一覧を返す。"""
    project = load_project(root)
    scene = pipeline.construct(project)
    return scene["slides"][0]["notes"], project.input_files


def test_imported_helper_is_tracked_and_reloaded_without_stale_bytecode(local_project):
    package = local_project / "local_content"
    package.mkdir()
    (package / "__init__.py").write_text("", encoding="utf-8")
    helper = package / "notes.py"
    helper.write_text('NOTE = "最初の内容を説明します。"\n', encoding="utf-8")
    fixed_time = 1700000000
    os.utime(helper, (fixed_time, fixed_time))
    source = local_project / "slides.py"
    source.write_text(
        "from local_content.notes import NOTE\n"
        + source.read_text(encoding="utf-8").replace('notes="確認内容を説明します。"', "notes=NOTE"),
        encoding="utf-8",
    )
    notes, inputs = constructed(local_project)
    assert notes == "最初の内容を説明します。" and inputs[str(helper)] == file_hash(helper)
    helper.write_text('NOTE = "次回の内容を説明します。"\n', encoding="utf-8")
    os.utime(helper, (fixed_time, fixed_time))
    notes, inputs = constructed(local_project)
    assert notes == "次回の内容を説明します。" and inputs[str(helper)] == file_hash(helper)
    assert not any(name == "local_content" or name.startswith("local_content.") for name in sys.modules)


def test_imported_helper_is_tracked_without_listing(local_project):
    # inputs.python を書かない資料でも、入口と、入口が読み込んだ補助のファイルを追跡する。
    assert "python" not in tomllib.loads((local_project / "deck.toml").read_text(encoding="utf-8"))["inputs"]
    helper = local_project / "deck_notes.py"
    helper.write_text('NOTE = "補助の内容を説明します。"\n', encoding="utf-8")
    source = local_project / "slides.py"
    source.write_text(
        "from deck_notes import NOTE\n"
        + source.read_text(encoding="utf-8").replace('notes="確認内容を説明します。"', "notes=NOTE"),
        encoding="utf-8",
    )
    notes, inputs = constructed(local_project)
    assert notes == "補助の内容を説明します。"
    assert {str(helper), str(source)} <= set(inputs)
    assert "deck_notes" not in sys.modules


def test_preloaded_shared_helper_is_fresh_for_build_and_restored(local_project, monkeypatch):
    import importlib

    shared = local_project.parent / "shared"
    shared.mkdir()
    helper = shared / "slidekit_preloaded_notes.py"
    helper.write_text('NOTE = "古い共有内容を説明します。"\n', encoding="utf-8")
    monkeypatch.syspath_prepend(str(shared))
    previous = importlib.import_module("slidekit_preloaded_notes")
    try:
        helper.write_text('NOTE = "現在の共有内容を説明します。"\n', encoding="utf-8")
        settings = local_project / "deck.toml"
        settings.write_text(
            settings.read_text(encoding="utf-8").replace('read_roots = ["."]', 'read_roots = [".", "../shared"]'),
            encoding="utf-8",
        )
        source = local_project / "slides.py"
        source.write_text(
            "from slidekit_preloaded_notes import NOTE\n"
            + source.read_text(encoding="utf-8").replace('notes="確認内容を説明します。"', "notes=NOTE"),
            encoding="utf-8",
        )
        notes, inputs = constructed(local_project)
        assert notes == "現在の共有内容を説明します。"
        assert inputs[str(helper)] == file_hash(helper)
        assert sys.modules["slidekit_preloaded_notes"] is previous
        assert previous.NOTE == "古い共有内容を説明します。"
    finally:
        sys.modules.pop("slidekit_preloaded_notes", None)


BIBLIOGRAPHY = """@book{knuth1997,
  author = {Donald E. Knuth},
  title = {The Art of Computer Programming},
  publisher = {Addison-Wesley},
  year = {1997},
}

@article{dijkstra1959,
  author = {Edsger W. Dijkstra},
  title = {A Note on Two Problems in Connexion with Graphs},
  journal = {Numerische Mathematik},
  year = {1959},
}
"""


def test_construct_records_bibliography_keys_and_entry_evidence(local_project):
    # 文献ファイルは 1 回だけ読み、キーの一覧と文献項目の証拠を設定に入れる。ページの内容の証拠は参照する項目でだけ変わる。
    references = local_project / "references.bib"
    references.write_text(BIBLIOGRAPHY, encoding="utf-8")
    settings = local_project / "deck.toml"
    settings.write_text(
        settings.read_text(encoding="utf-8").replace(
            'read_roots = ["."]', 'read_roots = ["."]\nbibliography = ["references.bib"]'
        ),
        encoding="utf-8",
    )
    source = local_project / "slides.py"
    source.write_text(
        source.read_text(encoding="utf-8").replace(
            "area=Rect(60, 110, 800, 90))", 'area=Rect(60, 110, 800, 90), sources=["knuth1997"])'
        ),
        encoding="utf-8",
    )

    def evidence():
        project = load_project(local_project)
        scene = pipeline.construct(project)
        return project, content_digest(scene["slides"][0], project)

    project, base = evidence()
    assert project.config["_bibliography_keys"] == ["dijkstra1959", "knuth1997"]
    assert project.config["_bibliography_entries"] == {
        key: project.bibliography.digest(key) for key in ("knuth1997", "dijkstra1959")
    }
    references.write_text(BIBLIOGRAPHY.replace("Numerische Mathematik", "Numer. Math."), encoding="utf-8")
    assert evidence()[1] == base
    references.write_text(BIBLIOGRAPHY.replace("Addison-Wesley", "Addison Wesley"), encoding="utf-8")
    assert evidence()[1] != base


def test_extended_deck_imports_parent_helper(tmp_path, write_deck, dejavu, dejavu_environment):
    # 継承した設定の入口は、自分のディレクトリの補助を読み込むことができる。同じ名前の補助は資料のディレクトリのものを使う。
    base = write_deck(
        tmp_path / "base",
        deck=DECK,
        theme=dejavu.theme,
        slides="""from slidekit import Deck, Rect
from deck_text import NOTE
from base_text import TITLE

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="intro", kind="blank", title=TITLE, notes=NOTE)
    slide.text(id="body", text="Process checked", area=Rect(60, 110, 800, 90))
    return deck
""",
    )
    (base / "deck_text.py").write_text('NOTE = "共通の説明です。"\n', encoding="utf-8")
    (base / "base_text.py").write_text('TITLE = "共通の題名"\n', encoding="utf-8")
    child = tmp_path / "child"
    child.mkdir()
    (child / "deck.toml").write_text('schema_version = 1\nextends = "../base/deck.toml"\n', encoding="utf-8")
    (child / "deck_text.py").write_text('NOTE = "この資料の説明です。"\n', encoding="utf-8")
    path_before = list(sys.path)
    project = load_project(child)
    scene = pipeline.construct(project)
    assert (scene["slides"][0]["title"], scene["slides"][0]["notes"]) == ("共通の題名", "この資料の説明です。")
    assert {str(base / "slides.py"), str(base / "base_text.py"), str(child / "deck_text.py")} <= set(
        project.input_files
    )
    assert str(base / "deck_text.py") not in project.input_files
    assert sys.path == path_before
    assert not {"deck_text", "base_text"} & set(sys.modules)


def test_page_number_format_and_frame_change_carries_visuals_as_renumbered(local_project):
    # ページ番号の書式と枠だけを変えると、番号の図形を除いた証拠が一致し、番号の領域（変更前と後の両方）を除いた画像も一致する。
    # 表示の確認はページ番号だけ変わったページとして引き継ぎ、内容の確認もそのまま引き継ぐ。
    rewrite(local_project, pages_source("first", "second"))
    first = built(local_project)
    finish_review(first)
    theme = local_project / "theme.toml"
    theme.write_text(
        theme.read_text(encoding="utf-8")
        + '[page_kinds.content]\npage_number_format = "{page} / {total}"\nfooter_area_pt = [820, 505, 100, 24]\n',
        encoding="utf-8",
    )
    result = pipeline.build(local_project / "deck.toml")
    second = Path(result["build"])
    assert result["carry"]["renumbered"] == ["first", "second"]
    copies = [record for record in records._review(second)["records"] if record["kind"] == "visual"]
    assert {record["slide_id"] for record in copies} == {"first", "second"}
    for record in copies:
        check = record["render_check"]
        assert check["footer_only"] is True and len(check["masked"]) == 2
        assert check["source_evidence_hash"] != record["evidence_hash"]
    status = records.review_status(second)
    pending = {group["kind"] for group in status["reviews"] if group["pending"]}
    # 一覧画像は毎回確かめ直す。ページの表示と内容は引き継いだ。
    assert pending == {"overview"}
    # 引き継いだ写しは、最終検査と同じ基準でも有効である。
    assert all(records._valid_record(second, record, {})[0] for record in copies)


def test_build_compares_page_images_with_previous_run_and_writes_sheets(local_project):
    # 生成の後に、前の実行とページ画像を比べる。変えたページだけを「変わった」とし、比較画像に枠を描く。
    rewrite(local_project, pages_source("first", "second", "third"))
    first = built(local_project)
    finish_review(first)
    rewrite(local_project, pages_source("first", "second", "third", titles={"second": "Changed title"}))
    result = pipeline.build(local_project / "deck.toml")
    second = Path(result["build"])
    comparison = result["comparison"]
    assert comparison["run"] == first.name
    assert (comparison["unchanged"], comparison["changed"], comparison["added"]) == (["first", "third"], ["second"], [])
    from slidekit.report import summarize

    assert (
        f"前の実行 {first.name} との比較: 変更なし 2、ページ番号だけ変わった 0、変わった 1（p.2）"
        in summarize("build", result).splitlines()
    )
    sheet = records.review_sheet(second)
    assert sheet["slides"] == ["second"] and sheet["changes"] == {"second": "changed"}
    assert sheet["compared_run"] == first.name
    assert all(Path(path).is_file() and Path(path).parent == second / "sheets" for path in sheet["sheets"])
    # 比較画像は成果物ではないので、整合検査に影響しない。
    assert records.check_build(second)["status"] == "passed"
