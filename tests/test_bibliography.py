import sys

import pytest

from slidekit import bibliography
from slidekit.bibliography import (
    Bibliography,
    BibliographyError,
    Reference,
    entry_digest,
    format_reference,
    load_bibliography,
)
from slidekit.errors import InputError
from slidekit.fragments import Math, Text, valid_link

# 試験に使う .bib は、公開されている一般的な文献だけにする。
CLASSICS = r"""
@article{erdos1959,
  author = {Erd{\H o}s, Paul and R{\'e}nyi, Alfr{\'e}d},
  title = {On random graphs {I}},
  journal = {Publicationes Mathematicae Debrecen},
  volume = 6, pages = {290--297}, year = 1959
}
@article{goedel1931,
  author = {G{\"o}del, Kurt},
  title = {{\"U}ber formal unentscheidbare S{\"a}tze der {Principia Mathematica} und verwandter {Systeme I}},
  journal = {Monatshefte f{\"u}r Mathematik und Physik},
  volume = 38, pages = {173--198}, year = 1931,
  note = {Reprinted in \textit{Collected Works}, Vol.~I}
}
@article{schoenhage1971,
  author = {Sch{\"o}nhage, Arnold and Strassen, Volker},
  title = {Schnelle {M}ultiplikation gro{\ss}er {Z}ahlen},
  journal = {Computing}, volume = 7, pages = {281--292}, year = 1971
}
@inproceedings{aks1983,
  author = {Mikl{\'o}s Ajtai and J{\'a}nos Koml{\'o}s and Endre Szemer{\'e}di},
  title = {An {$O(n \log n)$} sorting network},
  booktitle = {Proceedings of the 15th Annual {ACM} Symposium on Theory of Computing},
  pages = {1--9}, year = 1983
}
@inproceedings{legall2014,
  author = {Fran{\c{c}}ois {Le Gall}},
  title = {Powers of tensors and fast matrix multiplication},
  booktitle = {Proceedings of the 39th International Symposium on Symbolic and Algebraic Computation},
  pages = {296--303}, year = 2014,
  note = {Shows \(\omega < 2.3728639\)}
}
@book{brooks1995,
  author = {Brooks, Jr., Frederick P.},
  title = {The Mythical Man-Month},
  publisher = {Addison-Wesley}, year = 1995,
  note = {Anniversary edition---with four new chapters}
}
@inproceedings{veb1975,
  author = {van Emde Boas, Peter},
  title = {Preserving order in a forest in less than logarithmic time},
  booktitle = {16th Annual Symposium on Foundations of Computer Science},
  pages = {75--84}, year = 1975
}
@book{knuth1984book,
  author = {Donald~E. Knuth},
  title = {The {\TeX}book},
  publisher = {Addison-Wesley}, year = 1984
}
@book{knuth1968,
  author = {Knuth, Donald E.},
  title = {The {\em Art} of Computer Programming, Volume 1: Fundamental Algorithms},
  publisher = {Addison-Wesley}, year = 1968
}
@misc{knuthhome,
  author = {Donald E. Knuth},
  title = {Home page},
  howpublished = {https://www-cs-faculty.stanford.edu/\~{}knuth/}, year = 2024
}
@book{matousek2002,
  author = {Ji{\v{r}}{\'\i} Matou{\v{s}}ek},
  title = {Lectures on Discrete Geometry},
  publisher = {Springer}, year = 2002
}
@book{clrs1990,
  author = {Thomas H. Cormen and Charles E. Leiserson and Ronald L. Rivest},
  title = {Introduction to Algorithms},
  publisher = {MIT Press \& McGraw-Hill}, year = 1990
}
@book{martinlof1984,
  author = {Per Martin-L{\"o}f},
  title = {Intuitionistic Type Theory},
  publisher = {Bibliopolis}, year = 1984
}
@book{takagi1983,
  author = {高木, 貞治},
  title = {解析概論},
  publisher = {岩波書店}, year = 1983
}
@book{yoshida2013,
  author = {吉田 洋一 and 赤 攝也},
  title = {数学序説},
  publisher = {筑摩書房}, year = 2013
}
@misc{unicode,
  author = {{The Unicode Consortium}},
  title = {The {Unicode} Standard},
  howpublished = {\url{https://example.com/~unicode/a_b%20c--d}},
  note = {See \href{https://example.com/guide}{the \emph{guide} page}},
  year = 2024
}
"""

KNUTH = r"""
@article{knuth1984,
  author = {Donald E. Knuth},
  title = {Literate Programming},
  journal = {The Computer Journal},
  volume = {27}, number = {2}, pages = {97--111}, year = {1984},
  doi = {10.1093/comjnl/27.2.97}
}
"""


def write(directory, name, text):
    path = directory / name
    path.write_text(text, encoding="utf-8")
    return path


def replace(text, *pairs):
    """試験用の .bib の一部を書き換える。書き換える文字列がなければ失敗にする（書き換えないまま試験が通らないように）。"""
    for old, new in zip(pairs[::2], pairs[1::2]):
        assert old in text, old
        text = text.replace(old, new)
    return text


def test_latex_text_and_names_are_decoded(tmp_path):
    bib = load_bibliography([write(tmp_path, "classics.bib", CLASSICS)])
    assert bib["erdos1959"].authors == ("Paul Erdős", "Alfréd Rényi")
    assert bib["erdos1959"].fields["title"] == "On random graphs I"  # 波括弧を外す
    assert bib["erdos1959"].fields["pages"] == "290–297"
    goedel = bib["goedel1931"]
    assert goedel.type == "article" and goedel.authors == ("Kurt Gödel",)
    assert goedel.fields["title"] == (
        "Über formal unentscheidbare Sätze der Principia Mathematica und verwandter Systeme I"
    )
    assert goedel.fields["journal"] == "Monatshefte für Mathematik und Physik"
    assert goedel.fields["note"] == "Reprinted in Collected Works, Vol. I"  # 書式の命令は中身に、~ は空白に
    assert bib["schoenhage1971"].authors == ("Arnold Schönhage", "Volker Strassen")
    assert bib["schoenhage1971"].fields["title"] == "Schnelle Multiplikation großer Zahlen"
    assert bib["legall2014"].authors == ("François Le Gall",)
    assert bib["matousek2002"].authors == ("Jiří Matoušek",)
    assert bib["martinlof1984"].authors == ("Per Martin-Löf",)
    assert bib["knuth1984book"].authors == ("Donald E. Knuth",)
    assert bib["knuth1984book"].fields["title"] == "The TeXbook"
    assert bib["clrs1990"].fields["publisher"] == "MIT Press & McGraw-Hill"
    # 書体の切り替え（{\em ...}）は中身だけを残し、空の引数のアクセント（\~{}）は記号にする。
    assert bib["knuth1968"].fields["title"] == "The Art of Computer Programming, Volume 1: Fundamental Algorithms"
    assert bib["knuthhome"].fields["howpublished"] == "https://www-cs-faculty.stanford.edu/~knuth/"
    # von と Jr.：「名 von 姓, Jr.」の順にする。
    assert bib["veb1975"].authors == ("Peter van Emde Boas",)
    assert bib["brooks1995"].authors == ("Frederick P. Brooks, Jr.",)
    assert bib["brooks1995"].fields["note"] == "Anniversary edition—with four new chapters"
    # 和文の名前は「姓 名」。「姓, 名」の形は並べ替え、「姓 名」の形は書いた順のまま。
    assert bib["takagi1983"].authors == ("高木 貞治",)
    assert bib["yoshida2013"].authors == ("吉田 洋一", "赤 攝也")
    # 数式は $...$ で囲んで残し、断片は変換したときの区切りから作る。
    aks = bib["aks1983"]
    assert aks.authors == ("Miklós Ajtai", "János Komlós", "Endre Szemerédi")
    assert aks.fields["title"] == r"An $O(n \log n)$ sorting network"
    assert aks.fragments("title", bold=True) == [
        Text("An ", bold=True),
        Math(r"O(n \log n)", bold=True),
        Text(" sorting network", bold=True),
    ]
    assert bib["legall2014"].fragments("note") == [Text("Shows "), Math(r"\omega < 2.3728639")]
    # \url の URL はそのまま読み（% ~ -- _ を変換しない）、\url と \href はリンクの断片にする。
    unicode = bib["unicode"]
    assert unicode.authors == ("The Unicode Consortium",)
    url = "https://example.com/~unicode/a_b%20c--d"
    assert unicode.fields["howpublished"] == url
    assert unicode.fragments("howpublished") == [Text(url, link=url)]
    assert unicode.fields["note"] == "See the guide page"
    assert unicode.fragments("note") == [Text("See "), Text("the guide page", link="https://example.com/guide")]
    # 断片の書式は指定したものだけを使い、数式に対応していない書式や未知の書式は誤りにする。
    with pytest.raises(BibliographyError, match="underline・link"):
        aks.fragments("title", underline=True)
    with pytest.raises(TypeError, match="italic"):
        aks.fragments("title", italic=True)
    with pytest.raises(KeyError):
        aks.fragments("abstract")


def test_url_like_fields_keep_their_text(tmp_path):
    text = replace(
        KNUTH,
        "doi = {10.1093/comjnl/27.2.97}",
        r"doi = {10.1093/comjnl/27.2.97}, url = {https://example.com/a\_b\%20c--d~e/{x}}",
    )
    reference = load_bibliography([write(tmp_path, "knuth.bib", text)])["knuth1984"]
    # url・doi などは波括弧を外して記号のエスケープだけを戻し、-- や ~ を変換しない。
    assert reference.fields["url"] == "https://example.com/a_b%20c--d~e/x"
    assert reference.fields["doi"] == "10.1093/comjnl/27.2.97"


def test_supported_commands_are_not_dropped():
    # 対応する命令はどれも、知らないうちに消えない（\url・\href は別の試験で確かめる）。
    quiet = {" ", ",", "-", "/", "@", "em", "bf", "it", "sl", "sc", "rm", "tt", "sf"}  # 空白か、何も出さない命令
    for name in sorted(bibliography.SUPPORTED_COMMANDS - quiet - {"url", "href"}):
        if name in bibliography._ACCENT_MARKS:
            source, unexpected = f"\\{name}{{e}}", "e"
        elif name in (
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
        ):
            source, unexpected = f"\\{name}{{abc}}", None
        else:
            source, unexpected = f"\\{name}{{}}", ""
        converted = bibliography._plain("test", "note", source)
        assert converted and converted != unexpected, name
        if unexpected is None:
            assert converted == "abc", name
    assert bibliography._plain("test", "note", r"a\ b\,c\-d\/e\@f {\em g}") == "a b cdef g"


def test_unknown_latex_command_is_an_error(tmp_path):
    text = replace(
        KNUTH,
        "title = {Literate Programming}",
        r"title = {Literate \foo{Programming} \cite{x} $\log n$}",
        "journal = {The Computer Journal}",
        "journal = {The Computer Journal}, note = {100% pure}",
    )
    text += r"""
@article{dijkstra1959,
  author = {Dijkstra, Edsger W.},
  title = {A note on two problems in connexion with graphs},
  abstract = {Uses \begin{itemize}\item{x}\end{itemize}},
  journal = {Numerische Mathematik}, volume = 1, pages = {269--271}, year = 1959,
  annote = {Unclosed $x}
}
"""
    bib = load_bibliography([write(tmp_path, "knuth.bib", text)])
    # キーの一覧と証拠は変換しないので、読み込みでは誤りにならない。
    assert list(bib) == ["knuth1984", "dijkstra1959"] and bib.digest("knuth1984")
    knuth = bib["knuth1984"]
    assert knuth.fields["journal"] == "The Computer Journal"  # 項目ごとに変換する
    with pytest.raises(ValueError) as caught:
        knuth.fields["title"]
    assert (
        str(caught.value) == r"文献 knuth1984 の title に未対応の LaTeX 命令があります: \foo、\cite"
    )  # 数式の中は調べない
    assert isinstance(caught.value, InputError) and caught.value.code == "invalid_reference"
    assert caught.value.details["commands"] == [r"\foo", r"\cite"]
    with pytest.raises(BibliographyError, match="エスケープしていない %"):
        knuth.fields["note"]
    with pytest.raises(BibliographyError, match="未対応の LaTeX 命令"):
        format_reference(knuth)
    dijkstra = bib["dijkstra1959"]
    with pytest.raises(BibliographyError, match=r"\\begin\{itemize\}、\\item$"):
        dijkstra.fields["abstract"]
    with pytest.raises(BibliographyError, match="数式が閉じていません"):
        dijkstra.fields["annote"]
    # 一覧に使わない項目の誤りは、文献一覧の項目の文章を妨げない。
    assert format_reference(dijkstra)[0] == Text("Edsger W. Dijkstra\n")


def test_duplicate_keys_across_files_are_rejected(tmp_path):
    first = write(tmp_path, "first.bib", KNUTH)
    second = write(tmp_path, "second.bib", replace(KNUTH, "Literate Programming", "Literate programming"))
    with pytest.raises(ValueError) as caught:
        load_bibliography([first, second])
    assert str(caught.value) == f"文献の識別子が重複しています: knuth1984（{first}、{second}）"
    assert isinstance(caught.value, InputError) and caught.value.code == "invalid_reference"
    with pytest.raises(BibliographyError, match="文献の識別子が重複しています: knuth1984"):
        load_bibliography([write(tmp_path, "twice.bib", KNUTH + KNUTH)])
    with pytest.raises(BibliographyError, match="同じ項目が複数あります: title"):
        load_bibliography(
            [write(tmp_path, "fields.bib", replace(KNUTH, "year = {1984}", "year = {1984}, title = {X}"))]
        )
    with pytest.raises(BibliographyError, match="文献情報を解析できません"):
        load_bibliography([write(tmp_path, "broken.bib", "@article{knuth1984 title = {Literate Programming}}")])
    with pytest.raises(BibliographyError, match="文献の識別子が空です"):
        load_bibliography([write(tmp_path, "nokey.bib", "@article{, title = {Literate Programming}}")])
    # 別々のキーなら、ファイル順に並ぶ。
    other = write(tmp_path, "other.bib", CLASSICS)
    assert list(load_bibliography([first, other]))[:3] == ["knuth1984", "erdos1959", "goedel1931"]


def test_format_reference_by_entry_type(tmp_path):
    bib = load_bibliography([write(tmp_path, "knuth.bib", KNUTH), write(tmp_path, "classics.bib", CLASSICS)])
    doi = "https://doi.org/10.1093/comjnl/27.2.97"
    assert format_reference(bib["knuth1984"]) == [
        Text("Donald E. Knuth\n"),
        Text("Literate Programming", bold=True),
        Text("\nThe Computer Journal, 27(2), 97–111, 1984\n"),
        Text(doi, link=doi),
    ]
    assert format_reference(bib["aks1983"], max_authors=2) == [
        Text("Miklós Ajtai, János Komlós, et al.\n"),
        Text("An ", bold=True),
        Math(r"O(n \log n)", bold=True),
        Text(" sorting network", bold=True),
        Text("\nProceedings of the 15th Annual ACM Symposium on Theory of Computing, 1–9, 1983"),
    ]
    assert format_reference(bib["aks1983"])[0] == Text("Miklós Ajtai, János Komlós, Endre Szemerédi\n")
    assert format_reference(bib["brooks1995"]) == [
        Text("Frederick P. Brooks, Jr.\n"),
        Text("The Mythical Man-Month", bold=True),
        Text("\nAddison-Wesley, 1995"),
    ]
    url = "https://example.com/~unicode/a_b%20c--d"
    assert format_reference(bib["unicode"]) == [
        Text("The Unicode Consortium\n"),
        Text("The Unicode Standard", bold=True),
        Text("\n"),
        Text(url, link=url),
        Text(", 2024"),
    ]
    # 利用者が作った Reference も整形できる。著者の others は et al. になる。
    reference = Reference(
        key="note", type="misc", fields={"title": "Notes", "year": "2026"}, authors=("A. Author", "et al.")
    )
    assert format_reference(reference, max_authors=1) == [
        Text("A. Author, et al.\n"),
        Text("Notes", bold=True),
        Text("\n2026"),
    ]
    for value in (0, True, "2"):
        with pytest.raises(BibliographyError, match="max_authors"):
            format_reference(reference, max_authors=value)


def test_bibliography_entry_digest_is_per_entry(tmp_path):
    source = KNUTH + r"""
@proceedings{stoc1983,
  title = {Proceedings of the 15th Annual ACM Symposium on Theory of Computing},
  booktitle = {Proceedings of the 15th Annual ACM Symposium on Theory of Computing},
  publisher = {ACM}, year = 1983
}
@inproceedings{aks1983,
  author = {Mikl{\'o}s Ajtai and J{\'a}nos Koml{\'o}s and Endre Szemer{\'e}di},
  title = {An {$O(n \log n)$} sorting network},
  crossref = {stoc1983}, pages = {1--9}
}
"""
    keys = ("knuth1984", "stoc1983", "aks1983")

    def digests(text, name):
        bib = load_bibliography([write(tmp_path, name, text)])
        return {key: bib.digest(key) for key in keys}

    base = digests(source, "base.bib")
    assert len(set(base.values())) == 3
    # 書式（空白・改行・引用符と波括弧・項目の順・名前と種類の大文字小文字）では変わらない。
    reformatted = replace(
        source,
        "@article{knuth1984,",
        "@ARTICLE{knuth1984,",
        "title = {Literate Programming},",
        'TITLE   =   "Literate\n      Programming" ,',
        "volume = {27}, number = {2}",
        "number = 2,\n  volume = 27",
    )
    assert digests(reformatted, "reformatted.bib") == base
    # 値を変えた項目と、それを crossref で参照する項目だけが変わる。
    changed = digests(replace(source, "Literate Programming", "Literate programming"), "title.bib")
    assert changed["knuth1984"] != base["knuth1984"]
    assert changed["stoc1983"] == base["stoc1983"] and changed["aks1983"] == base["aks1983"]
    changed = digests(replace(source, "publisher = {ACM}", "publisher = {ACM Press}"), "parent.bib")
    assert changed["stoc1983"] != base["stoc1983"] and changed["aks1983"] != base["aks1983"]
    assert changed["knuth1984"] == base["knuth1984"]
    # 証拠は bibtexparser が読んだままの値（変換の前）で作る。参照先の値は fields に引き継ぐ。
    bib = load_bibliography([tmp_path / "base.bib"])
    assert bib.digest("knuth1984") == entry_digest(
        "Article",
        "knuth1984",
        {
            "Author": "Donald E. Knuth",
            "title": "Literate Programming",
            "journal": "The Computer Journal",
            "volume": "27",
            "number": "2",
            "pages": "97--111",
            "year": "1984",
            "doi": "10.1093/comjnl/27.2.97",
        },
        None,
    )
    assert bib["aks1983"].fields["publisher"] == "ACM"
    # 参照先がない、または循環する crossref でも止まらない。
    looped = digests(replace(source, "crossref = {stoc1983}", "crossref = {aks1983}"), "loop.bib")
    missing = digests(replace(source, "crossref = {stoc1983}", "crossref = {stoc1984}"), "missing.bib")
    assert len({base["aks1983"], looped["aks1983"], missing["aks1983"]}) == 3


def test_bibliography_mapping_is_lazy_and_reports_unknown_keys(tmp_path):
    path = write(tmp_path, "knuth.bib", KNUTH)
    bib = load_bibliography(path)  # パス 1 つでもよい
    assert len(bib) == 1 and "knuth1984" in bib and "knuth1948" not in bib and bib.paths == (path,)
    with pytest.raises(KeyError) as caught:
        bib["knuth1948"]
    assert str(caught.value) == "未定義の文献キーです: knuth1948（候補: knuth1984）"
    assert isinstance(caught.value, InputError) and caught.value.code == "invalid_reference"
    with pytest.raises(KeyError, match="未定義の文献キーです: dijkstra1959$"):
        bib.digest("dijkstra1959")
    assert bib.get("knuth1948") is None and bib["knuth1984"] is bib["knuth1984"]
    empty = load_bibliography([])
    assert isinstance(empty, Bibliography) and len(empty) == 0 and list(Bibliography()) == []


def test_missing_bibtexparser_has_guidance(tmp_path, monkeypatch):
    path = write(tmp_path, "knuth.bib", KNUTH)
    with monkeypatch.context() as patch:
        patch.setitem(sys.modules, "bibtexparser", None)
        with pytest.raises(ImportError) as caught:
            load_bibliography([path])
        assert str(caught.value) == (
            "文献ファイルの読み込みには追加依存 bibliography（bibtexparser）が必要です。"
            f"uv pip install -e '.[bibliography]' で導入してください: {path}"
        )
        assert len(load_bibliography([])) == 0  # 文献を使わない資料は追加依存なしで動く
    bib = load_bibliography([path])
    bibliography._latex.cache_clear()
    try:
        with monkeypatch.context() as patch:
            patch.setitem(sys.modules, "pylatexenc", None)
            with pytest.raises(
                ImportError, match=r"bibliography（pylatexenc）.*uv pip install -e '\.\[bibliography\]'"
            ):
                bib["knuth1984"]
    finally:
        bibliography._latex.cache_clear()
    assert bib["knuth1984"].authors == ("Donald E. Knuth",)


def test_links_follow_fragment_link_rule(tmp_path):
    # 文献のリンクは断片のリンクと同じ形式（fragments.valid_link：https://・http://・mailto: で始まり空白を含まない）だけにする。
    text = r"""
@misc{archive,
  title = {Data archive},
  url = {ftp://example.org/data},
  note = {Mirror at \href{ftp://example.org/mirror}{the mirror}, contact \url{mailto:archive@example.org}},
  year = 2020
}
"""
    reference = load_bibliography([write(tmp_path, "archive.bib", text)])["archive"]
    last = format_reference(reference)[-1]
    assert last.text.endswith("\nftp://example.org/data") and last.link is None
    assert reference.fragments("note") == [
        Text("Mirror at the mirror, contact "),
        Text("mailto:archive@example.org", link="mailto:archive@example.org"),
    ]
    assert valid_link("mailto:archive@example.org") and not valid_link("ftp://example.org/data")
