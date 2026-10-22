"""試験の共通の準備。利用者の環境設定と Windows のフォントを読まないようにする。

すべての試験で、空の環境設定（SLIDEKIT_ENVIRONMENT）を使い、既定フォントの自動探索を止め、SLIDEKIT_JOBS を 1 にする。
試験の中で設定し直せば、そちらが優先される。資料の設定を書く関数と、DejaVu Sans の準備、表示ソフトの偽の変換セッション
（fake_session）もここに置く。
"""

import contextlib
import json
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from slidekit import fonts

# 自動探索の本来の関数。探索そのものを確かめる試験と、実際の Windows のフォントを使う試験だけが使う。
REAL_DETECT_FONT_DIRECTORIES = fonts._detect_font_directories

# 資料の設定の最小形。
DECK = """schema_version = 1
entrypoint = "slides.py:build"
[metadata]
title = "試験の資料"
author = ""
last_modified_by = ""
"""

# 何もしない生成関数。
EMPTY_BUILD = "def build(context):\n    pass\n"

DEJAVU_DIRECTORY = Path("/usr/share/fonts/truetype/dejavu")


def toml_string(value) -> str:
    """TOML の基本文字列として書く（JSON の文字列の表記は TOML でも読むことができる）。"""
    return json.dumps(str(value))


def pytest_configure(config):
    if os.environ.get("SLIDEKIT_TEST_POWERPOINT") == "1" and getattr(config.option, "numprocesses", None):
        raise pytest.UsageError("SLIDEKIT_TEST_POWERPOINT=1 の試験は直列に実行する。-n を外すか -p no:xdist を付ける")


@pytest.fixture(autouse=True)
def isolated_environment(tmp_path_factory, monkeypatch):
    """空の環境設定を使い、既定フォントの自動探索を止める。環境設定のパスを返す。

    一時ファイルは tmp_path と別のフォルダーに作る（tmp_path の中身の数を確かめる試験に影響させない）。
    """
    path = tmp_path_factory.mktemp("environment") / "environment.toml"
    path.write_text("schema_version = 1\n", encoding="utf-8")
    monkeypatch.setenv("SLIDEKIT_ENVIRONMENT", str(path))
    # 並列の試験だけが上書きする。
    monkeypatch.setenv("SLIDEKIT_JOBS", "1")
    monkeypatch.setattr(fonts, "_detect_font_directories", lambda environment: ())
    # 自動探索の記憶を試験ごとに分ける（模擬の結果が後の試験へ残らないようにする）。
    monkeypatch.setattr(fonts, "_DETECTED", {})
    return path


@pytest.fixture
def font_detection(isolated_environment, monkeypatch):
    """既定フォントの自動探索を本来の関数に戻し、その関数を返す。探索そのものを確かめる試験が使う。"""
    monkeypatch.setattr(fonts, "_detect_font_directories", REAL_DETECT_FONT_DIRECTORIES)
    return REAL_DETECT_FONT_DIRECTORIES


@pytest.fixture
def write_deck():
    """deck.toml と slides.py を書いて資料のディレクトリを返す関数。

    theme を渡すと theme.toml も書き、deck.toml の先頭から参照する（このとき deck には theme を書かない）。
    extra は deck.toml の末尾に加える。
    """

    def write(root, *, deck=DECK, slides=EMPTY_BUILD, theme=None, extra=""):
        root = Path(root)
        root.mkdir(parents=True, exist_ok=True)
        if theme is not None:
            (root / "theme.toml").write_text(theme, encoding="utf-8")
            deck = 'theme = "theme.toml"\n' + deck
        (root / "deck.toml").write_text(deck + extra, encoding="utf-8")
        (root / "slides.py").write_text(slides, encoding="utf-8")
        return root

    return write


@pytest.fixture
def dejavu():
    """検査用の DejaVu Sans。通常と太字のファイル、すべての書体を DejaVu Sans にするテーマ、その実体を指定する環境設定の文。"""
    path, bold_path = DEJAVU_DIRECTORY / "DejaVuSans.ttf", DEJAVU_DIRECTORY / "DejaVuSans-Bold.ttf"
    if not path.is_file() or not bold_path.is_file():
        pytest.skip("検査用の DejaVu Sans がない")
    family = "DejaVu Sans"
    theme = f"[fonts]\njapanese = {toml_string(family)}\nlatin = {toml_string(family)}\nmath = {toml_string(family)}\n"
    environment = f"schema_version = 1\n[fonts.{toml_string(family)}]\npath = {toml_string(path)}\nbold_path = {toml_string(bold_path)}\n"
    return SimpleNamespace(family=family, path=path, bold_path=bold_path, theme=theme, environment=environment)


@pytest.fixture
def dejavu_environment(isolated_environment, dejavu):
    """共通の環境設定で DejaVu Sans の実体を指定する。環境設定のパスを返す。"""
    isolated_environment.write_text(dejavu.environment, encoding="utf-8")
    return isolated_environment


@pytest.fixture
def windows_fonts_environment(isolated_environment):
    """実際の Windows のフォントのフォルダーを指定した環境設定にする。既定の 3 書体がなければ省略する。

    資料を構築する PowerPoint の試験と、使用例を構築する試験が使う。PowerShell でフォルダーを求めるのは
    SLIDEKIT_TEST_POWERPOINT=1 のときだけで、ほかでは PowerShell を使わずに求めたフォルダーを使う。
    """
    if os.environ.get("SLIDEKIT_TEST_POWERPOINT") == "1":
        folders = REAL_DETECT_FONT_DIRECTORIES({})
    else:
        folders = fonts._existing(fonts._fixed_font_directories())
    environment = {"font_directories": [str(folder) for folder in folders]}
    if any(fonts.default_font(family, environment) is None for family in ("Yu Gothic UI", "Arial", "Cambria Math")):
        pytest.skip("Windows の既定フォントがない")
    listing = ", ".join(toml_string(folder) for folder in environment["font_directories"])
    isolated_environment.write_text(f"schema_version = 1\nfont_directories = [{listing}]\n", encoding="utf-8")
    return isolated_environment


# 偽の変換が読む OOXML の名前空間。
_OOXML = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _qualified(prefix, name):
    return f"{{{_OOXML[prefix]}}}{name}"


def _offset(node):
    """図形の位置と寸法（pt）。p:spPr か p:grpSpPr の a:xfrm から読む。"""
    transform = node.find("p:spPr/a:xfrm", _OOXML)
    if transform is None:
        transform = node.find("p:grpSpPr/a:xfrm", _OOXML)
    if transform is None:
        return None
    offset, extent = transform.find("a:off", _OOXML), transform.find("a:ext", _OOXML)
    if offset is None or extent is None:
        return None
    return (
        int(offset.get("x")) / 12700,
        int(offset.get("y")) / 12700,
        int(extent.get("cx")) / 12700,
        int(extent.get("cy")) / 12700,
    )


def _paragraph_text(paragraph):
    """段落の文字（a:t と数式の m:t を文書の順に）と、最初の a:rPr の大きさ（pt。なければ 20）。"""
    text = "".join(node.text or "" for node in paragraph.iter(_qualified("a", "t"), _qualified("m", "t")))
    # 数式だけの段落の幅 0 の空白（U+200B）は描かない。
    text = text.replace("\u200b", "")
    sizes = [node.get("sz") for node in paragraph.iter(_qualified("a", "rPr")) if node.get("sz")]
    return text, int(sizes[0]) / 100 if sizes else 20


def _draw_shape(page, node, part, font):
    """偽の変換で 1 つの図形を描く。PowerPoint と同じく、mc:AlternateContent は mc:Choice（編集可能な数式）を描く。"""
    if node.tag == _qualified("mc", "AlternateContent"):
        choice = node.find("mc:Choice", _OOXML)
        for child in choice if choice is not None else ():
            _draw_shape(page, child, part, font)
        return
    if node.tag == _qualified("p", "grpSp"):
        for child in node:
            _draw_shape(page, child, part, font)
        return
    if node.tag not in (_qualified("p", "sp"), _qualified("p", "pic")):
        return
    area = _offset(node)
    if area is None:
        return
    left, top, width, height = area
    blip = node.find(".//a:blip", _OOXML)
    if blip is not None and blip.get(_qualified("r", "embed")):
        # 画像と、互換画像を塗りにした図形（確認用複製）。
        image = part.related_part(blip.get(_qualified("r", "embed")))
        page.insert_image((left, top, left + width, top + height), stream=image.blob, keep_proportion=False)
    body = node.find("p:txBody", _OOXML)
    if body is None:
        return
    for number, paragraph in enumerate(body.findall("a:p", _OOXML)):
        text, size = _paragraph_text(paragraph)
        if text.strip():
            page.insert_text(
                (left, top + size * 1.2 * (number + 1)), text, fontfile=str(font), fontname="DejaVuSans", fontsize=size
            )


def render_fixture_pdf(input_pptx, output_pdf, *, font):
    """保存した PPTX の文字（数式の文字を含む）と画像を読み、検査用の PDF を作る。"""
    from pdf_tools import Document
    from pptx import Presentation

    presentation = Presentation(input_pptx)
    output_pdf = Path(output_pdf)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with Document() as document:
        for slide in presentation.slides:
            page = document.new_page(width=presentation.slide_width.pt, height=presentation.slide_height.pt)
            for node in slide.element.find("p:cSld/p:spTree", _OOXML):
                _draw_shape(page, node, slide.part, font)
        document.save(output_pdf)
    return output_pdf


class FakeSession:
    """`renderers.open_session()` の偽物が返すセッション。事前確認は固定の値、変換は render_fixture_pdf で行う。"""

    def __init__(self, calls, name, font):
        self.calls, self.name, self.font = calls, name, font
        self.renders = 0
        self.closed = False

    def doctor(self):
        self.calls.doctors.append(self.name)
        if self.calls.on_doctor is not None:
            self.calls.on_doctor(self.name)
        return {
            "ok": True,
            "renderer": self.name,
            "version": "fixture-1",
            "executable": f"{self.name}-fixture",
            "windows_user_sid": "S-1-5-21-fixture",
            "font_matches": [],
        }

    def render(self, input_pptx, output_pdf, *, label="render"):
        from slidekit.files import file_hash

        input_pptx, output_pdf = Path(input_pptx), Path(output_pdf)
        if output_pdf.exists():
            raise FileExistsError(f"PDF 出力先には新しいパスを指定する: {output_pdf}")
        self.calls.renders.append(
            {"renderer": self.name, "label": label, "input": input_pptx, "input_sha256": file_hash(input_pptx)}
        )
        if self.calls.on_render is not None:
            self.calls.on_render(self.name, label, input_pptx, output_pdf)
        render_fixture_pdf(input_pptx, output_pdf, font=self.font)
        self.renders += 1
        return {
            "renderer": self.name,
            "version": "fixture-1",
            "build": "1",
            "executable": f"{self.name}-fixture",
            "label": label,
            "elapsed_seconds": 0.0,
            "input_sha256": file_hash(input_pptx),
            "output_sha256": file_hash(output_pdf),
        }

    def close(self):
        if not self.closed:
            self.closed = True
            self.calls.closed.append(self.name)
        return {"renderer": self.name, "closed": True}

    @property
    def operations(self):
        return {"powershell_launches": 0, "renders": self.renders, "waited_seconds": 0.0}


@pytest.fixture
def fake_session(monkeypatch, dejavu):
    """生成の流れの表示ソフトを偽の変換セッションに差し替え、呼び出しを記録する。

    戻り値の sessions に開いたセッション（名前と引数）、doctors に事前確認、renders に変換（表示ソフト・label・入力）、
    closed に閉じたセッションが入る。on_open・on_doctor・on_render に関数を入れると、その時点で呼ぶ（例外を出せば失敗を再現できる）。
    """
    from slidekit import pipeline

    calls = SimpleNamespace(
        sessions=[], doctors=[], renders=[], closed=[], on_open=None, on_doctor=None, on_render=None
    )

    @contextlib.contextmanager
    def open_session(name, environment, *, wait_seconds=None, holder=None, on_event=None, state_dir=None):
        calls.sessions.append({"name": name, "wait_seconds": wait_seconds, "holder": holder})
        if calls.on_open is not None:
            calls.on_open(name, on_event)
        session = FakeSession(calls, name, dejavu.path)
        try:
            yield session
        finally:
            session.close()

    monkeypatch.setattr(pipeline, "open_session", open_session)
    return calls
