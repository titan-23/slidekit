"""文献ファイル（.bib）の読み込み、LaTeX 表記の変換、著者名の整形、文献一覧の項目の文章。

読み込みには追加依存 bibliography（bibtexparser と、その依存の pylatexenc）を使う。
LaTeX 表記は pylatexenc で変換し、対応する命令（`SUPPORTED_COMMANDS`）にない命令は消さずに誤りにする。
変換は項目を参照したときに行う。キーの一覧と文献項目の証拠（`Bibliography.digest`）は変換しない。
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator, Mapping
import dataclasses
from dataclasses import dataclass
import difflib
from functools import cache
from itertools import islice
import os
from pathlib import Path
import re
from types import SimpleNamespace
import unicodedata

from .errors import ArgumentNameError, InputError
from .files import json_digest
from .fragments import Math, Text, valid_link

GUIDANCE = (
    "文献ファイルの読み込みには追加依存 bibliography（bibtexparser）が必要です。"
    "uv pip install -e '.[bibliography]' で導入してください: {path}"
)
LATEX_GUIDANCE = (
    "文献の LaTeX 表記の変換には追加依存 bibliography（pylatexenc）が必要です。"
    "uv pip install -e '.[bibliography]' で導入してください"
)

# LaTeX 表記を変換せず、波括弧を外して記号のエスケープだけを戻す項目。
VERBATIM_FIELDS = frozenset({"url", "doi", "eprint", "isbn", "issn"})

# 空の引数のアクセント（例：\~{}）は、アクセントの記号そのものにする。
_ACCENT_MARKS = {
    "'": "´",
    "`": "`",
    "^": "^",
    '"': "¨",
    "~": "~",
    "=": "¯",
    ".": "˙",
    "u": "˘",
    "v": "ˇ",
    "H": "˝",
    "c": "¸",
    "k": "˛",
    "r": "˚",
    "d": ".",
    "b": "_",
}
_ESCAPES = tuple("&%$#_{}")
_LETTERS = (
    "i",
    "j",
    "ss",
    "o",
    "O",
    "ae",
    "AE",
    "oe",
    "OE",
    "aa",
    "AA",
    "l",
    "L",
    "DH",
    "dh",
    "TH",
    "th",
    "DJ",
    "dj",
    "NG",
    "ng",
)
_LOGOS = ("TeX", "LaTeX", "BibTeX")
_SYMBOLS = (
    "textendash",
    "textemdash",
    "textquoteleft",
    "textquoteright",
    "textquotedblleft",
    "textquotedblright",
    "textbackslash",
    "textasciitilde",
    "textasciicircum",
    "textdollar",
    "textregistered",
    "textcopyright",
    "texttrademark",
    "textdegree",
    "textsection",
    "textparagraph",
    "textellipsis",
    "ldots",
    "dots",
    "textbullet",
    "dag",
    "textexclamdown",
    "textquestiondown",
    "guillemotleft",
    "guillemotright",
    "textsterling",
    "texteuro",
)
_SPACING = (" ", ",", "-", "/", "@")  # 空白（\ と \,）と、表示に影響しない命令（\- \/ \@）
_CONTENT = (
    "emph",
    "textit",
    "textbf",
    "textsc",
    "textrm",
    "texttt",
    "textsf",
    "textsl",
    "textup",
    "textmd",
    "textnormal",
    "mbox",
    "text",
)  # 中身だけを残す（書体の指定は使わない）
_SWITCHES = ("em", "bf", "it", "sl", "sc", "rm", "tt", "sf")  # {\em ...} などの書体の切り替え（何も出さない）
SUPPORTED_COMMANDS = frozenset(
    [*_ESCAPES, *_ACCENT_MARKS, *_LETTERS, *_LOGOS, *_SYMBOLS, *_SPACING, *_CONTENT, *_SWITCHES, "url", "href"]
)

_SPACE = re.compile(r"[ \t\r\n\f\v]+")  # LaTeX が 1 つの空白とみなす文字（全角の空白は含めない）


class BibliographyError(InputError):
    """文献ファイルと文献項目の誤り。`code` は `invalid_reference`。"""

    code = "invalid_reference"


class BibliographyKeyError(BibliographyError, KeyError):
    """未定義の文献キー。`KeyError` としても捕まえることができる。"""

    def __str__(self):
        return str(self.args[0]) if self.args else ""


def _collapse(value: str) -> str:
    return _SPACE.sub(" ", value).strip()


def _verbatim(value: str) -> str:
    """URL などの値から、エスケープしていない波括弧を除き、記号のエスケープ（\\% など）を戻す。"""
    result, index = [], 0
    while index < len(value):
        character = value[index]
        if character == "\\" and index + 1 < len(value) and value[index + 1] in "&%$#_{}~^":
            result.append(value[index + 1])
            index += 2
            continue
        if character not in "{}":
            result.append(character)
        index += 1
    return _collapse("".join(result))


@cache
def _latex():
    """1 度だけ pylatexenc を読み込み、文献用に直した変換の設定を作る。"""
    try:
        from pylatexenc import latex2text, latexwalker, macrospec
    except ImportError as error:
        raise ImportError(LATEX_GUIDANCE) from error

    class LinkArguments(macrospec.MacroStandardArgsParser):
        """\\url・\\href の URL を、波括弧の対応だけを見てそのまま読む（% や ~ や -- を変換しない）。"""

        def parse_args(self, w, pos, parsing_state=None):
            if parsing_state is None:
                parsing_state = w.make_parsing_state()
            start = pos
            while start < len(w.s) and w.s[start].isspace():
                start += 1
            if start >= len(w.s) or w.s[start] != "{":
                raise latexwalker.LatexWalkerParseError(msg="URL を波括弧で囲んでください", s=w.s, pos=start)
            depth, index = 0, start
            while index < len(w.s):
                character = w.s[index]
                if character == "\\":
                    index += 2
                    continue
                if character == "{":
                    depth += 1
                elif character == "}":
                    depth -= 1
                    if depth == 0:
                        break
                index += 1
            else:
                raise latexwalker.LatexWalkerParseError(msg="URL の { が閉じていません", s=w.s, pos=start)
            url = w.make_node(
                latexwalker.LatexCharsNode,
                parsing_state=parsing_state,
                chars=w.s[start + 1 : index],
                pos=start + 1,
                len=index - start - 1,
            )
            arguments, end = [url], index + 1
            if self.argspec == "{{":
                node, node_pos, node_len = w.get_latex_expression(end, strict_braces=False, parsing_state=parsing_state)
                arguments.append(node)
                end = node_pos + node_len
            return macrospec.ParsedMacroArgs(argspec=self.argspec, argnlist=arguments), pos, end - pos

    class Converter(latex2text.LatexNodes2Text):
        """数式と URL の位置に目印を入れて変換し、文字と数式とリンクの区切りを残す。"""

        def __init__(self, context, source):
            super().__init__(latex_context=context, math_mode="verbatim", strict_latex_spaces=True)
            unused = (chr(code) for code in range(0xE000, 0xF900) if chr(code) not in source)
            self.open, self.close = islice(unused, 2)
            self.inserts = []

        def insert(self, kind, value, link=None):
            self.inserts.append((kind, value, link))
            return f"{self.open}{len(self.inserts) - 1}{self.close}"

        def math_node_to_text(self, node):
            start, end = node.delimiters
            verbatim = node.latex_verbatim()
            return self.insert("math", verbatim[len(start) : len(verbatim) - len(end)])

    def url_text(node, l2tobj):
        url = _verbatim(node.nodeargd.argnlist[0].chars)
        return l2tobj.insert("text", url, url if valid_link(url) else None)

    def href_text(node, l2tobj):
        url = _verbatim(node.nodeargd.argnlist[0].chars)
        text = l2tobj.node_arg_to_text(node, 1)
        if not valid_link(url):
            return text
        return l2tobj.insert("link", url) + text + l2tobj.insert("end", None)

    def accent(original, mark):
        def convert(node, l2tobj):
            argument = node.nodeargd.argnlist[0] if node.nodeargd and node.nodeargd.argnlist else None
            if not l2tobj.nodelist_to_text([argument]).strip():
                return mark
            return original(node, l2tobj)

        return convert

    walker = latexwalker.get_default_latex_context_db()
    walker.add_context_category(
        "slidekit-bibliography",
        prepend=True,
        macros=[
            macrospec.MacroSpec("url", args_parser=LinkArguments("{")),
            macrospec.MacroSpec("href", args_parser=LinkArguments("{{")),
            macrospec.MacroSpec(
                "textnormal", args_parser=macrospec.MacroStandardArgsParser("{", args_math_mode=[False])
            ),
            macrospec.MacroSpec("\\", ""),  # 改行（対応しない）。末尾の \\ でも引数を探さずに命令として報告する
        ],
    )
    text = latex2text.get_default_latex_context_db()
    Spec = latex2text.MacroTextSpec
    macros = [
        Spec(name, simplify_repl=accent(text.get_macro_spec(name).simplify_repl, mark))
        for name, mark in _ACCENT_MARKS.items()
    ]
    macros += [Spec(name, discard=False) for name in ("texttt", "textsf", "textup", "textmd", "textnormal", "mbox")]
    macros += [Spec(name, simplify_repl=name) for name in _LOGOS]
    macros += [Spec(name, discard=True) for name in ("-", "/", "@", *_SWITCHES)]
    macros += [
        Spec("textasciicircum", simplify_repl="^"),
        Spec("url", simplify_repl=url_text),
        Spec("href", simplify_repl=href_text),
    ]
    text.add_context_category(
        "slidekit-bibliography",
        prepend=True,
        macros=macros,
        specials=[latex2text.SpecialsTextSpec("~", " "), latex2text.SpecialsTextSpec("&", "&")],
    )
    return SimpleNamespace(latexwalker=latexwalker, walker=walker, text=text, Converter=Converter)


def _commands(latex, nodes, found: dict, comments: list) -> None:
    """対応していない命令の名前と、エスケープしていない % を集める（数式の中は調べない）。"""
    lw = latex.latexwalker
    for node in nodes:
        if node is None or node.isNodeType(lw.LatexMathNode):
            continue
        if node.isNodeType(lw.LatexCommentNode):
            comments.append(node)
        elif node.isNodeType(lw.LatexEnvironmentNode):
            found.setdefault(f"\\begin{{{node.environmentname}}}")
            _commands(latex, node.nodelist, found, comments)
        elif node.isNodeType(lw.LatexGroupNode):
            _commands(latex, node.nodelist, found, comments)
        elif node.isNodeType(lw.LatexMacroNode) or node.isNodeType(lw.LatexSpecialsNode):
            arguments = list(node.nodeargd.argnlist) if node.nodeargd and node.nodeargd.argnlist else []
            if node.isNodeType(lw.LatexMacroNode):
                if node.macroname not in SUPPORTED_COMMANDS:
                    found.setdefault("\\" + node.macroname)
                if node.macroname in ("url", "href"):
                    arguments = arguments[1:]  # URL はそのまま読んだ文字
            _commands(latex, arguments, found, comments)


def _parse_problem(error) -> str:
    message = getattr(error, "msg", None) or str(error)
    missing = re.search(r'End of input while parsing arguments of macro "(.*)"', message)
    if "was expecting '$'" in message or "was expecting '\\)'" in message or "was expecting '\\]'" in message:
        message = "数式が閉じていません"
    elif "was expecting '}'" in message:
        message = "{ が閉じていません"
    elif "mismatching closing brace" in message:
        message = "対応する { のない } があります"
    elif missing:
        message = f"\\{missing.group(1)} の引数がありません"
    position = getattr(error, "pos", None)
    return message if position is None else f"{message}（{position + 1} 文字目）"


def _segments(key: str, name: str, value: str) -> tuple[tuple[str, str, str | None], ...]:
    """項目の値を（種類、文字、リンク）の列にする。種類は text か math。"""
    if name in VERBATIM_FIELDS:
        text = _verbatim(value)
        return (("text", text, None),) if text else ()
    latex = _latex()
    try:
        nodes, start, length = latex.latexwalker.LatexWalker(
            value, latex_context=latex.walker, tolerant_parsing=False
        ).get_latex_nodes()
        if start + length < len(value):
            raise latex.latexwalker.LatexWalkerParseError(msg="解析できない文字があります", s=value, pos=start + length)
    except latex.latexwalker.LatexWalkerError as error:
        raise BibliographyError(
            f"文献 {key} の {name} の LaTeX 表記を解析できません: {_parse_problem(error)}", key=key, field=name
        ) from error
    found, comments = {}, []
    _commands(latex, nodes, found, comments)
    if found:
        commands = list(found)
        raise BibliographyError(
            f"文献 {key} の {name} に未対応の LaTeX 命令があります: {'、'.join(commands)}",
            key=key,
            field=name,
            commands=commands,
        )
    if comments:
        raise BibliographyError(
            f"文献 {key} の {name} にエスケープしていない % があります（文字の % は \\% と書きます）",
            key=key,
            field=name,
        )
    converter = latex.Converter(latex.text, value)
    converted = converter.nodelist_to_text(nodes)
    marker = re.compile(re.escape(converter.open) + r"(\d+)" + re.escape(converter.close))
    pieces, links = [], []
    for position, part in enumerate(marker.split(converted)):
        if position % 2 == 0:
            if converter.open in part or converter.close in part:
                raise BibliographyError(
                    f"文献 {key} の {name} の LaTeX 表記を解析できません: "
                    "数式と URL は、アクセントなどの命令の引数にできません",
                    key=key,
                    field=name,
                )
            if part:
                pieces.append(("text", part, links[-1] if links else None))
            continue
        kind, text, link = converter.inserts[int(part)]
        if kind == "link":
            links.append(text)
        elif kind == "end":
            links.pop()
        else:
            pieces.append((kind, text, link or (links[-1] if links and kind == "text" else None)))
    return _normalize(key, name, pieces)


def _normalize(key, name, pieces):
    """空白の並びを 1 つにし、前後の空白を除き、隣り合う同じリンクの文字をつなげる。"""
    result = []
    for kind, text, link in pieces:
        if kind == "math":
            text = " ".join(text.split())
            if not text:
                raise BibliographyError(f"文献 {key} の {name} の数式が空です", key=key, field=name)
            result.append((kind, text, None))
            continue
        text = _SPACE.sub(" ", text)
        if not result or (result[-1][0] == "text" and result[-1][1].endswith(" ")):
            text = text.lstrip(" ")  # 先頭の空白と、空白に続く空白を除く
        if not text:
            continue
        if result and result[-1][0] == "text" and result[-1][2] == link:
            result[-1] = ("text", result[-1][1] + text, link)
        else:
            result.append(("text", text, link))
    if result and result[-1][0] == "text":
        last = result[-1][1].rstrip(" ")
        if last:
            result[-1] = ("text", last, result[-1][2])
        else:
            result.pop()
    return tuple(result)


def _plain(key: str, name: str, value: str) -> str:
    """表示用の文字列（数式は $...$ で囲む）。"""
    return "".join(text if kind == "text" else f"${text}$" for kind, text, _ in _segments(key, name, value))


def _top_level_comma(name: str) -> bool:
    depth, index = 0, 0
    while index < len(name):
        character = name[index]
        if character == "\\":
            index += 2
            continue
        if character == "{":
            depth += 1
        elif character == "}":
            depth = max(0, depth - 1)
        elif character == "," and depth == 0:
            return True
        index += 1
    return False


def _names(key: str, name: str, value: str | None) -> tuple[str, ...]:
    """著者・編者の表示用の名前。和文字を含む名前は「姓 名」、それ以外は「名 von 姓, Jr.」にする。"""
    if not value:
        return ()
    from bibtexparser.middlewares.names import parse_single_name_into_parts, split_multiple_persons_names

    result = []
    for person in split_multiple_persons_names(value):
        if person.strip().lower() == "others":
            result.append("et al.")
            continue
        parts = parse_single_name_into_parts(person, strict=False)
        first, von, last, jr = (
            _plain(key, name, " ".join(words)) for words in (parts.first, parts.von, parts.last, parts.jr)
        )
        wide = any(unicodedata.east_asian_width(character) in "WF" for character in first + von + last + jr)
        # 和文の名前を「姓, 名」と書いたときだけ並べ替える。「姓 名」と書いた名前は書いた順のまま。
        words = (von, last, first) if wide and _top_level_comma(person) else (first, von, last)
        display = " ".join(word for word in words if word)
        result.append(f"{display}, {jr}" if jr else display)
    return tuple(person for person in result if person)


class _Fields(Mapping):
    """LaTeX 表記を直した表示用の値。各項目は最初に参照したときに変換する。"""

    def __init__(self, key: str, raw: Mapping[str, str]):
        self._key, self._raw, self._segments = key, dict(raw), {}

    def segments(self, name: str) -> tuple[tuple[str, str, str | None], ...]:
        if name not in self._raw:
            raise KeyError(name)
        if name not in self._segments:
            self._segments[name] = _segments(self._key, name, self._raw[name])
        return self._segments[name]

    def __getitem__(self, name: str) -> str:
        return "".join(text if kind == "text" else f"${text}$" for kind, text, _ in self.segments(name))

    def __iter__(self) -> Iterator[str]:
        return iter(self._raw)

    def __len__(self) -> int:
        return len(self._raw)

    def __contains__(self, name) -> bool:
        return name in self._raw

    def __repr__(self) -> str:
        return f"<文献 {self._key} の項目: {'、'.join(self._raw)}>"


_STYLE_NAMES = frozenset({"bold", "color", "size_pt", "underline", "superscript", "link"})
_MATH_STYLE_NAMES = frozenset({"bold", "color", "size_pt"})


@dataclass(frozen=True)
class Reference:
    """1 件の文献。`fields` は LaTeX 表記を直した表示用の値（名前は小文字。数式は $...$ で囲む）。"""

    key: str
    type: str  # article、inproceedings、book など（小文字）
    fields: Mapping[str, str] = dataclasses.field(hash=False)
    authors: tuple[str, ...] = ()  # 表示用の名前
    editors: tuple[str, ...] = ()

    def fragments(self, field: str, **style) -> list[Text | Math]:
        """項目を文章の断片にする。変換したときの文字と数式とリンクの区切りを使い、fields を解析し直さない。

        `style` は各断片の書式（`bold`・`color`・`size_pt`・`underline`・`link`）。数式には `bold`・`color`・`size_pt` を使う。
        """
        unknown = sorted(set(style) - _STYLE_NAMES)
        if unknown:
            raise ArgumentNameError(
                f"fragments() に未知の書式があります: {'、'.join(unknown)}（使える書式: "
                f"{'、'.join(sorted(_STYLE_NAMES))}）"
            )
        fields = self.fields
        segments = fields.segments(field) if isinstance(fields, _Fields) else (("text", fields[field], None),)
        result = []
        for kind, text, link in segments:
            if kind == "math":
                if set(style) - _MATH_STYLE_NAMES:
                    raise BibliographyError(
                        f"文献 {self.key} の {field} は数式を含むので underline・link を指定できません",
                        key=self.key,
                        field=field,
                    )
                result.append(Math(text, **style))
            else:
                result.append(Text(text, **{"link": link, **style}))
        return result


class Bibliography(Mapping):
    """文献キーから `Reference` を引く。並びはファイル順。未定義のキーは候補付きの KeyError にする。"""

    def __init__(self, entries: Iterable = (), *, paths: Iterable = ()):
        self._entries = {entry.key: entry for entry in entries}
        self._references: dict[str, Reference] = {}
        self.paths = tuple(Path(path) for path in paths)

    def _missing(self, key) -> BibliographyKeyError:
        message = f"未定義の文献キーです: {key}"
        if isinstance(key, str):
            candidates = difflib.get_close_matches(key, list(self._entries), n=3, cutoff=0.6)
            if candidates:
                message += f"（候補: {'、'.join(candidates)}）"
        return BibliographyKeyError(message, key=key)

    def __getitem__(self, key: str) -> Reference:
        if key not in self._entries:
            raise self._missing(key)
        if key not in self._references:
            entry = self._entries[key]
            fields = dict(entry.fields)
            parent = self._entries.get(fields.get("crossref", "").strip())
            if parent is not None and parent is not entry:  # BibTeX と同じく、ない項目を参照先から 1 段だけ引き継ぐ
                for name, value in parent.fields.items():
                    if name != "crossref":
                        fields.setdefault(name, value)
            self._references[key] = Reference(
                key=key,
                type=entry.type,
                fields=_Fields(key, fields),
                authors=_names(key, "author", fields.get("author")),
                editors=_names(key, "editor", fields.get("editor")),
            )
        return self._references[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._entries)

    def __len__(self) -> int:
        return len(self._entries)

    def __contains__(self, key) -> bool:
        return key in self._entries

    def __repr__(self) -> str:
        return f"Bibliography({len(self)} 件)"

    def digest(self, key: str) -> str:
        """文献項目の証拠。bibtexparser が読んだままの値（変換の前。@string は展開済み）で作る。"""
        if key not in self._entries:
            raise self._missing(key)
        return self._digest(key, frozenset())

    def _digest(self, key: str, visiting: frozenset) -> str:
        entry = self._entries[key]
        chain = visiting | {key}
        target = entry.fields.get("crossref")
        if target is None:
            crossref = None
        elif target.strip() in chain:
            crossref = "cycle"
        elif target.strip() not in self._entries:
            crossref = "missing"
        else:
            crossref = self._digest(target.strip(), chain)
        return entry_digest(entry.type, key, entry.fields, crossref)


@dataclass(frozen=True)
class _Entry:
    key: str
    type: str
    fields: dict  # 名前は小文字。値は bibtexparser が読んだまま
    path: Path


def entry_digest(entry_type: str, key: str, fields: Mapping[str, str], crossref: str | None) -> str:
    """文献項目の証拠（正規化した JSON の SHA-256）。

    項目の名前は小文字にして名前順に並べ、値の空白の並びは 1 つにする。ファイルの書式（空白・改行・
    引用符と波括弧の違い・項目の順）では変わらず、値の変更で変わる。`crossref` は参照先の証拠
    （参照先がなければ "missing"、循環すれば "cycle"）。
    """
    items = sorted([str(name).lower(), _collapse(str(value))] for name, value in fields.items())
    return json_digest({"type": str(entry_type).lower(), "key": key, "fields": items, "crossref": crossref})


def _failure(path: Path, block) -> BibliographyError:
    from bibtexparser.model import DuplicateBlockKeyBlock, DuplicateFieldKeyBlock, Entry

    line = None if block.start_line is None else block.start_line + 1
    where = f"{path} の {line} 行目" if line else str(path)
    details = {"path": str(path), "line": line}
    if isinstance(block, DuplicateBlockKeyBlock):
        if isinstance(block.ignore_error_block, Entry):
            return BibliographyError(f"文献の識別子が重複しています: {block.key}（{where}）", key=block.key, **details)
        return BibliographyError(f"文献ファイルの @string が重複しています: {block.key}（{where}）", **details)
    if isinstance(block, DuplicateFieldKeyBlock):
        names = "、".join(sorted(block.duplicate_keys))
        key = block.ignore_error_block.key
        return BibliographyError(f"文献 {key} に同じ項目が複数あります: {names}（{where}）", key=key, **details)
    reason = getattr(block.error, "abort_reason", None) or str(block.error)
    return BibliographyError(
        f"文献情報を解析できません: {where}" + (f"（{reason.strip()}）" if reason.strip() else ""), **details
    )


def load_bibliography(paths) -> Bibliography:
    """文献ファイルを読む。解析できない項目、同じキー（ファイルをまたぐものを含む）、同じ項目名を誤りにする。

    `paths` が空なら bibtexparser を読み込まずに空の `Bibliography` を返す。
    """
    if isinstance(paths, (str, os.PathLike)):
        paths = [paths]
    paths = [Path(path) for path in paths]
    if not paths:
        return Bibliography()
    try:
        import bibtexparser
    except ImportError as error:
        raise ImportError(GUIDANCE.format(path=paths[0])) from error
    entries, owners = [], {}
    for path in paths:
        try:
            library = bibtexparser.parse_file(str(path))
        except UnicodeDecodeError as error:
            raise BibliographyError(f"文献ファイルを UTF-8 として読めません: {path}", path=str(path)) from error
        if library.failed_blocks:
            raise _failure(path, library.failed_blocks[0])
        for entry in library.entries:
            if not entry.key.strip():
                line = "" if entry.start_line is None else f" の {entry.start_line + 1} 行目"
                raise BibliographyError(f"文献の識別子が空です（{path}{line}）", path=str(path))
            if entry.key in owners:
                raise BibliographyError(
                    f"文献の識別子が重複しています: {entry.key}（{owners[entry.key]}、{path}）",
                    key=entry.key,
                    path=str(path),
                )
            owners[entry.key] = path
            fields = {}
            for item in entry.fields:
                name = item.key.lower()
                if name in fields:
                    raise BibliographyError(
                        f"文献 {entry.key} に同じ項目が複数あります: {name}（{path}）", key=entry.key, path=str(path)
                    )
                fields[name] = item.value if isinstance(item.value, str) else str(item.value)
            entries.append(_Entry(entry.key, entry.entry_type.lower(), fields, path))
    return Bibliography(entries, paths=paths)


def _field(reference: Reference, name: str, **style) -> list[Text | Math]:
    if name not in reference.fields or not reference.fields[name]:
        return []
    return reference.fragments(name, **style)


def _volume(reference: Reference) -> list[Text | Math]:
    volume, number = _field(reference, "volume"), _field(reference, "number")
    if number:
        number = [Text("("), *number, Text(")")]
    return volume + number


# 掲載情報に並べる項目（種類にないものは howpublished・publisher・organization の最初のものと年）。
_VENUES = {
    "article": ("journal", "volume", "pages", "year"),
    "inproceedings": ("booktitle", "pages", "year"),
    "incollection": ("booktitle", "pages", "year"),
    "conference": ("booktitle", "pages", "year"),
    "book": ("publisher", "year"),
    "inbook": ("publisher", "year"),
    "phdthesis": ("school", "year"),
    "mastersthesis": ("school", "year"),
    "techreport": ("institution", "number", "year"),
}


def _venue(reference: Reference) -> list[Text | Math]:
    names = _VENUES.get(reference.type)
    if names is None:
        first = next((name for name in ("howpublished", "publisher", "organization") if _field(reference, name)), None)
        names = (first, "year") if first else ("year",)
    pieces = [(_volume(reference) if name == "volume" else _field(reference, name)) for name in names]
    return _join([piece for piece in pieces if piece], ", ")


def _url(reference: Reference) -> str | None:
    url = reference.fields.get("url")
    if url:
        return url
    doi = reference.fields.get("doi")
    if not doi:
        return None
    doi = re.sub(r"^doi:\s*", "", doi, flags=re.IGNORECASE)
    return doi if re.match(r"https?://", doi, flags=re.IGNORECASE) else "https://doi.org/" + doi


def _join(parts: list[list], separator: str) -> list:
    result = []
    for number, part in enumerate(parts):
        if number:
            result.append(Text(separator))
        result.extend(part)
    return result


def _merge(fragments: list) -> list:
    """隣り合う同じ書式の Text をつなげる。"""
    result = []
    for fragment in fragments:
        previous = result[-1] if result else None
        if (
            type(fragment) is Text
            and type(previous) is Text
            and all(getattr(previous, name) == getattr(fragment, name) for name in _STYLE_NAMES)
        ):
            result[-1] = dataclasses.replace(previous, text=previous.text + fragment.text)
        else:
            result.append(fragment)
    return result


def format_reference(reference: Reference, *, max_authors: int | None = None) -> list[Text | Math]:
    """文献一覧の 1 項目の文章。著者、題名（太字）、掲載情報、URL の行を改行でつなぐ。ない項目の行は省く。

    `max_authors` を超える著者は、先頭から `max_authors` 人と `et al.` にする。
    """
    if max_authors is not None and (
        isinstance(max_authors, bool) or not isinstance(max_authors, int) or max_authors < 1
    ):
        raise BibliographyError(f"max_authors は 1 以上の整数で指定してください: {max_authors!r:.40}")
    lines = []
    authors = list(reference.authors)
    others = bool(authors) and authors[-1] == "et al."
    if others:
        authors.pop()
    if max_authors is not None and len(authors) > max_authors:
        authors, others = authors[:max_authors], True
    if authors or others:
        lines.append([Text(", ".join(authors + (["et al."] if others else [])))])
    title = _field(reference, "title", bold=True)
    if title:
        lines.append(title)
    venue = _venue(reference)
    if venue:
        lines.append(venue)
    url = _url(reference)
    if url:
        lines.append([Text(url, link=url) if valid_link(url) else Text(url)])
    return _merge(_join(lines, "\n"))
