"""テーマ・記述規則・deck.toml の設定の検査と、役割の使い方の判定を、辞書と設定ファイルで確かめる。"""

import copy
from pathlib import Path
import tomllib

import pytest

import slidekit
from slidekit.config import merge
from slidekit.errors import InputError
from slidekit.schema import (
    DECK_KEYS,
    DECK_LOCAL_KEYS,
    ENVIRONMENT_KEYS,
    MATH_CANDIDATE_KEYS,
    PAGE_KINDS,
    RULES_KEYS,
    Problem,
    RoleProblem,
    SettingsError,
    check_decorations,
    check_deck_requirements,
    check_keys,
    check_rules_references,
    check_settings,
    check_theme_references,
    role_usage_problems,
)

PACKAGE = Path(slidekit.__file__).resolve().parent
EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def read(path):
    with Path(path).open("rb") as stream:
        return tomllib.load(stream)


BUNDLED_THEME = read(PACKAGE / "themes/default.toml")
BUNDLED_RULES = read(PACKAGE / "rules/general.toml")


def problems_of(kind, data):
    """1 つのファイルの検査と、同梱の設定と合成した後の参照の検査を合わせた問題の文字列。"""
    if kind == "theme":
        own = check_settings("theme", data, file="theme.toml")
        return [str(p) for p in own + check_theme_references(merge(BUNDLED_THEME, data), file="theme.toml")]
    own = check_settings("rules", data, file="rules.toml")
    return [str(p) for p in own + check_rules_references(merge(BUNDLED_RULES, data), BUNDLED_THEME, file="rules.toml")]


def example_settings():
    for deck in sorted(EXAMPLES.glob("*/src/deck.toml")):
        config = read(deck)
        for kind, default in (("theme", "default"), ("rules", "general")):
            value = config.get(kind, default)
            if value != default:
                yield deck.parent.name, kind, deck.parent / value


def test_bundled_and_example_settings_are_valid():
    assert check_settings("theme", BUNDLED_THEME, file="同梱の themes/default.toml") == []
    assert check_settings("rules", BUNDLED_RULES, file="同梱の rules/general.toml") == []
    assert check_theme_references(BUNDLED_THEME, file="同梱の themes/default.toml") == []
    assert check_rules_references(BUNDLED_RULES, BUNDLED_THEME, file="同梱の rules/general.toml") == []
    decks = sorted(EXAMPLES.glob("*/src/deck.toml"))
    assert decks, "使用例の deck.toml が見つからない"
    for deck in decks:
        config = read(deck)
        assert check_settings("deck", config, file=str(deck)) == [], deck
        assert check_deck_requirements(config, file=str(deck)) == [], deck
    found = list(example_settings())
    assert found, "使用例の設定ファイルが見つからない"
    themes = {name: merge(BUNDLED_THEME, read(path)) for name, kind, path in found if kind == "theme"}
    for name, kind, path in found:
        data = read(path)
        assert check_settings(kind, data, file=str(path)) == [], path
        if kind == "theme":
            assert check_theme_references(themes[name], file=str(path)) == [], path
        else:
            assert (
                check_rules_references(merge(BUNDLED_RULES, data), themes.get(name, BUNDLED_THEME), file=str(path))
                == []
            ), path


def test_bundled_rules_state_every_default():
    assert set(BUNDLED_RULES) == set(RULES_KEYS)
    assert set(BUNDLED_RULES["math_candidates"]) == set(MATH_CANDIDATE_KEYS)
    assert BUNDLED_RULES["math_candidates"] == {
        "enabled": False,
        "numbers": True,
        "variables": [],
        "operations": [],
        "exclude_patterns": [],
        "exclude_roles": [],
        "exclude_page_kinds": [],
    }
    assert {key: BUNDLED_RULES[key] for key in ("forbidden_characters", "require_notes", "notes_style")} == {
        "forbidden_characters": [],
        "require_notes": False,
        "notes_style": "unrestricted",
    }


def test_kind_defaults_and_equation_role_pass_the_checks():
    # 種類別の既定と数式の役割を同梱テーマに加えても検査に通る。
    theme = merge(
        BUNDLED_THEME,
        {
            "defaults": {
                "box": {"fill": "surface", "line_color": "line", "padding_pt": 6},
                "table_cell": {"fill": "background"},
                "table_header": {"fill": "surface"},
                "equation": {"word_wrap": False},
                "arrow": {"line_color": "primary"},
            },
            "roles": {"equation": {"inherits": "body"}, "callout": {"inherits": "body", "fill": "FFF4D6"}},
        },
    )
    assert check_settings("theme", theme, file="theme.toml") == []
    assert check_theme_references(theme, file="theme.toml") == []


@pytest.mark.parametrize(
    ("kind", "data", "expected"),
    [
        (
            "rules",
            {"forbidden_character": ["。"]},
            "rules.toml: forbidden_character は不明な設定です（候補: forbidden_characters）",
        ),
        (
            "rules",
            {"math_candidates": {"exclude_pattern": []}},
            "rules.toml: math_candidates.exclude_pattern は不明な設定です（候補: exclude_patterns、exclude_page_kinds、exclude_roles）",
        ),
        ("theme", {"roles": {"note": {"size": 14}}}, "theme.toml: roles.note.size は不明な設定です（候補: size_pt）"),
        (
            "theme",
            {"page_kinds": {"contents": {"title_role": "slide_title"}}},
            "theme.toml: page_kinds.contents は不明なページ種別です（候補: content）",
        ),
        (
            "theme",
            {"page_kinds": {"content": {"title_styles": {}}}},
            "theme.toml: page_kinds.content.title_styles は不明な設定です（候補: title_style、title_role、footer_style）",
        ),
        ("theme", {"colours": {}}, "theme.toml: colours は不明な設定です（候補: colors）"),
        (
            "theme",
            {"defaults": {"boxes": {"fill": "surface"}}},
            "theme.toml: defaults.boxes は不明な設定です（候補: box）",
        ),
        (
            "theme",
            {"defaults": {"box": {"colour": "text"}}},
            "theme.toml: defaults.box.colour は不明な設定です（候補: color）",
        ),
        ("theme", {"defaults": {"line": {"fill": "surface"}}}, "theme.toml: defaults.line.fill は line では使えません"),
        (
            "theme",
            {"defaults": {"z_index": 3}},
            "theme.toml: defaults.z_index には対応していません。描画順は作成順で、後に作った要素が前面になります",
        ),
        ("theme", {"fonts": {"japanes": "Meiryo"}}, "theme.toml: fonts.japanes は不明な設定です（候補: japanese）"),
    ],
)
def test_unknown_setting_names_name_file_key_and_candidates(kind, data, expected):
    assert problems_of(kind, data) == [expected]


@pytest.mark.parametrize(
    ("kind", "data", "expected"),
    [
        (
            "rules",
            {"math_candidates": {"enabled": "yes"}},
            "rules.toml: math_candidates.enabled は真偽値で指定してください（値: 'yes'）",
        ),
        (
            "rules",
            {"math_candidates": {"exclude_patterns": ["ok", "("]}},
            "rules.toml: math_candidates.exclude_patterns の 2 番目は正規表現として解釈できません（missing ), unterminated subpattern at position 0）",
        ),
        (
            "rules",
            {"notes_style": "desu_masu"},
            "rules.toml: notes_style は unrestricted、polite のどれかで指定してください（値: 'desu_masu'）",
        ),
        (
            "rules",
            {"forbidden_characters": ["。", "です"]},
            "rules.toml: forbidden_characters の 2 番目は 1 文字で指定してください（値: 'です'）",
        ),
        (
            "rules",
            {"forbidden_characters": "。"},
            "rules.toml: forbidden_characters は配列で指定してください（値: '。'）",
        ),
        (
            "rules",
            {"math_candidates": {"variables": ["n", "time step"]}},
            "rules.toml: math_candidates.variables の 2 番目は空白を含まない空でない文字列で指定してください（値: 'time step'）",
        ),
        (
            "rules",
            {"math_candidates": {"exclude_page_kinds": ["reference"]}},
            "rules.toml: math_candidates.exclude_page_kinds の 1 番目は cover、section、content、references、blank のどれかで指定してください"
            "（値: 'reference'）（候補: references）",
        ),
        ("rules", {"require_notes": 1}, "rules.toml: require_notes は真偽値で指定してください（値: 1）"),
        ("rules", {"schema_version": 2}, "rules.toml: schema_version は 1 で指定してください（値: 2）"),
        (
            "theme",
            {"colors": {"teal": "07F87"}},
            "theme.toml: colors.teal は 6 桁の 16 進数で指定してください（値: '07F87'）",
        ),
        (
            "theme",
            {"colors": {"none": "FFFFFF"}},
            "theme.toml: colors.none は色の名前に使えません（none は塗らない・線なしの指定に使います）",
        ),
        (
            "theme",
            {"colors": {"light-blue": "DDEEFF"}},
            "theme.toml: colors.light-blue は色の名前に使えません（英字か _ で始め、英字・数字・_ だけを使います）",
        ),
        (
            "theme",
            {"page_kinds": {"cover": {"title_area": [0.1, 0.5, 0.95, 0.2]}}},
            "theme.toml: page_kinds.cover.title_area は 0 以上 1 以下の比で、幅と高さを正、x + 幅と y + 高さを 1 以下にしてください"
            "（値: [0.1, 0.5, 0.95, 0.2]）",
        ),
        (
            "theme",
            {"page_kinds": {"content": {"footer_area_pt": [600, 500, 0, 20]}}},
            "theme.toml: page_kinds.content.footer_area_pt の幅と高さは正にしてください（値: [600, 500, 0, 20]）",
        ),
        (
            "theme",
            {"page_kinds": {"content": {"page_number": "yes"}}},
            "theme.toml: page_kinds.content.page_number は真偽値で指定してください（値: 'yes'）",
        ),
        (
            "theme",
            {"page_kinds": {"content": {"page_number_format": "{pages} / {total}"}}},
            "theme.toml: page_kinds.content.page_number_format は {page} を含み、波括弧の欄が {page} と {total} だけの"
            "文字列で指定してください（値: '{pages} / {total}'）",
        ),
        (
            "theme",
            {"page_kinds": {"content": {"page_number_format": "{page} / {count}"}}},
            "theme.toml: page_kinds.content.page_number_format は {page} を含み、波括弧の欄が {page} と {total} だけの"
            "文字列で指定してください（値: '{page} / {count}'）",
        ),
        (
            "theme",
            {"page_kinds": {"content": {"title_style": {"align": "middle"}}}},
            "theme.toml: page_kinds.content.title_style.align は left、center、right、justify のどれかで指定してください（値: 'middle'）",
        ),
        (
            "theme",
            {"margins": {"left": -4}},
            "theme.toml: margins.left は 0 以上の有限の数で指定してください（値: -4）",
        ),
        (
            "theme",
            {"roles": {"caption": {"inherits": "note", "allowed_content": "image"}}},
            "theme.toml: roles.caption.allowed_content は math、text のどれかで指定してください（値: 'image'）",
        ),
        (
            "theme",
            {"roles": {"caption": {"inherits": "note", "restricted": "true"}}},
            "theme.toml: roles.caption.restricted は真偽値で指定してください（値: 'true'）",
        ),
        (
            "theme",
            {"roles": {"caption": {"inherits": "note", "allowed_slide_ids": ["intro", ""]}}},
            "theme.toml: roles.caption.allowed_slide_ids の 2 番目は空でない文字列で指定してください（値: ''）",
        ),
        (
            "theme",
            {"roles": {"caption-small": {"inherits": "note"}}},
            "theme.toml: roles.caption-small は役割の名前に使えません（英字か _ で始め、英字・数字・_ だけを使います）",
        ),
        (
            "theme",
            {"roles": {"note": {"padding_pt": [2, 4]}}},
            "theme.toml: roles.note.padding_pt は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください（値: [2, 4]）",
        ),
        (
            "theme",
            {"defaults": {"word_wrap": "no"}},
            "theme.toml: defaults.word_wrap は真偽値で指定してください（値: 'no'）",
        ),
    ],
)
def test_invalid_setting_values_are_rejected(kind, data, expected):
    assert problems_of(kind, data) == [expected]


@pytest.mark.parametrize(
    ("data", "expected"),
    [
        (
            {"roles": {"note": {"color": "mutd"}}},
            ["theme.toml: roles.note.color の色 mutd はテーマにありません（候補: muted）"],
        ),
        (
            {"defaults": {"box": {"fill": "tela"}}, "colors": {"teal": "007F87"}},
            ["theme.toml: defaults.box.fill の色 tela はテーマにありません（候補: teal）"],
        ),
        (
            {"roles": {"note": {"color": "none"}}},
            [
                "theme.toml: roles.note.color には none を指定できません（none は fill・line_color とページ種別の background だけに使えます）"
            ],
        ),
        (
            {"page_kinds": {"content": {"title_role": "slide_titles"}}},
            [
                "theme.toml: page_kinds.content.title_role の役割 slide_titles はテーマにありません（候補: slide_title、section_title、cover_title）"
            ],
        ),
        (
            {"roles": {"caption": {"color": "muted"}}},
            ["theme.toml: roles.caption に size_pt と min_size_pt を指定するか、inherits で継承してください"],
        ),
        (
            {"roles": {"a": {"inherits": "b"}, "b": {"inherits": "a"}}},
            ["theme.toml: roles.a.inherits の継承が循環しています（a → b → a）"],
        ),
        (
            {"roles": {"caption": {"inherits": "notes"}}},
            ["theme.toml: roles.caption.inherits の役割 notes はテーマにありません（候補: note）"],
        ),
        (
            {"roles": {"caption": {"inherits": "note", "size_pt": 10}}},
            ["theme.toml: roles.caption の size_pt は min_size_pt 以上にしてください（size_pt: 10、min_size_pt: 12）"],
        ),
    ],
)
def test_theme_references_are_checked_after_merge(data, expected):
    assert check_settings("theme", data, file="theme.toml") == []
    assert [str(p) for p in check_theme_references(merge(BUNDLED_THEME, data), file="theme.toml")] == expected


def test_reference_check_skips_values_reported_by_the_value_check():
    # 形の間違った色は値の検査だけが報告し、その種類の要素が受け付けない書式の色も参照の検査で二度報告しない。
    data = {"defaults": {"line": {"fill": "unknown"}}, "roles": {"note": {"color": "#12"}}}
    assert problems_of("theme", data) == [
        "theme.toml: defaults.line.fill は line では使えません",
        "theme.toml: roles.note.color は色の名前か 6 桁の 16 進数で指定してください（値: '#12'）",
    ]


def test_page_background_and_paint_accept_none():
    data = {
        "page_kinds": {"content": {"background": "none"}, "section": {"background": "surface"}},
        "defaults": {"box": {"line_color": "none"}},
        "roles": {"note": {"fill": "none"}},
    }
    assert problems_of("theme", data) == []
    assert problems_of("theme", {"page_kinds": {"cover": {"background": "paper"}}}) == [
        "theme.toml: page_kinds.cover.background の色 paper はテーマにありません"
    ]
    assert problems_of("theme", {"page_kinds": {"cover": {"background": 3}}}) == [
        "theme.toml: page_kinds.cover.background は色の名前か 6 桁の 16 進数で指定してください（値: 3）"
    ]
    assert problems_of("theme", {"defaults": {"emphasis_color": "none"}}) == [
        "theme.toml: defaults.emphasis_color には none を指定できません（none は fill・line_color とページ種別の background だけに使えます）"
    ]


def test_rules_exclude_roles_must_exist():
    data = {"math_candidates": {"exclude_roles": ["note", "referenses"], "exclude_page_kinds": ["references"]}}
    assert check_settings("rules", data, file="rules.toml") == []
    assert [str(p) for p in check_rules_references(merge(BUNDLED_RULES, data), BUNDLED_THEME, file="rules.toml")] == [
        "rules.toml: math_candidates.exclude_roles の 2 番目の役割 referenses はテーマにありません（候補: references）"
    ]
    theme = merge(BUNDLED_THEME, {"roles": {"caption": {"inherits": "note"}}})
    assert check_rules_references({"math_candidates": {"exclude_roles": ["caption"]}}, theme, file="rules.toml") == []
    assert problems_of("rules", {"math_candidates": {"exclude_roles": [""]}}) == [
        "rules.toml: math_candidates.exclude_roles の 1 番目は空でない文字列で指定してください（値: ''）"
    ]


def test_settings_problems_are_reported_together():
    problems = check_settings(
        "rules", {"forbidden_character": ["。"], "notes_style": "desu_masu"}, file="rules.toml"
    ) + check_settings("theme", {"roles": {"note": {"size": 14}}}, file="theme.toml")
    error = SettingsError(problems)
    assert isinstance(error, InputError) and isinstance(error, ValueError)
    assert error.code == "settings_invalid"
    assert str(error).splitlines() == [
        "rules.toml: forbidden_character は不明な設定です（候補: forbidden_characters）",
        "rules.toml: notes_style は unrestricted、polite のどれかで指定してください（値: 'desu_masu'）",
        "theme.toml: roles.note.size は不明な設定です（候補: size_pt）",
    ]
    assert error.problems == tuple(problems)
    assert error.details["problems"][2] == {
        "file": "theme.toml",
        "key": "roles.note.size",
        "message": "は不明な設定です",
        "suggestions": ("size_pt",),
    }


def test_problem_text_and_long_values():
    assert str(Problem("deck.toml", "render.requird", "は不明な設定です", ("required",))) == (
        "deck.toml: render.requird は不明な設定です（候補: required）"
    )
    assert (
        str(Problem("theme.toml", "", "TOML として読めません（Invalid value）"))
        == "theme.toml: TOML として読めません（Invalid value）"
    )
    [problem] = check_settings("rules", {"notes_style": "x" * 60}, file="rules.toml")
    shown = problem.message.split("（値: ")[1].removesuffix("）")
    assert len(shown) == 40 and shown.endswith("…")
    assert [str(p) for p in check_settings("theme", [], file="theme.toml")] == [
        "theme.toml: 設定は表形式で指定してください（値: []）"
    ]
    assert [str(p) for p in check_settings("theme", {"roles": {"note": 3}}, file="theme.toml")] == [
        "theme.toml: roles.note は表形式で指定してください（値: 3）"
    ]
    with pytest.raises(ValueError, match="未対応の設定の種類です: styles"):
        check_settings("styles", {}, file="styles.toml")


def test_check_keys_reports_unknown_deck_settings():
    assert [
        str(p)
        for p in check_keys(
            {"requird": ["powerpoint"], "primary": "powerpoint"}, {"primary", "required"}, "render", file="deck.toml"
        )
    ] == ["deck.toml: render.requird は不明な設定です（候補: required）"]
    assert [
        str(p) for p in check_keys({"entrypont": "slides.py:build"}, {"entrypoint", "size"}, "", file="deck.toml")
    ] == ["deck.toml: entrypont は不明な設定です（候補: entrypoint）"]
    assert [str(p) for p in check_keys(["x"], {"path"}, "template", file="deck.toml")] == [
        "deck.toml: template は表形式で指定してください（値: ['x']）"
    ]


# すべての設定名に正しい値を書いた deck.toml と実行環境。どの名前にも値の検査があることを確かめる。
FULL_DECK = {
    "schema_version": 1,
    "extends": "../base",
    "entrypoint": "slides.py:build",
    "size": [800, 600],
    "theme": "theme.toml",
    "rules": "general",
    "output_dir": "out",
    "environment_override": "local/environment.toml",
    "metadata": {name: "slidekit example" for name in DECK_KEYS["metadata"]},
    "inputs": {
        "read_roots": ["."],
        "python": ["helper.py"],
        "bibliography": ["references.bib"],
        "asset_dirs": ["figures"],
        "documents": [{"id": "outline", "path": "outline.md", "affects_content": True}],
    },
    "render": {"primary": "powerpoint", "required": ["powerpoint", "libreoffice"]},
    "math": {"fallback": "required", "fallback_renderer": "powerpoint", "fallback_scale": 2.5, "functions": ["argmax"]},
    "images": {"raster_scale": 4, "vector": True},
    "deliverables": {"pptx": "deck.pptx", "pdf": "deck.pdf", "notes": "notes.md", "record": "deliverables.json"},
    "template": {"path": "template.pptx", "layout_index": 6, "show_master_shapes": False},
}
FULL_ENVIRONMENT = {
    "schema_version": 1,
    "jobs": 4,
    "font_directories": ["fonts"],
    "fonts": {"Yu Gothic UI": {"path": "YuGothM.ttc", "index": 1, "bold_path": "YuGothB.ttc", "bold_index": 1}},
    "renderers": {
        "powerpoint": {"powershell": "powershell.exe", "timeout": 90, "wait_seconds": 0},
        "libreoffice": {"executable": "soffice", "timeout": 60.5},
    },
}


def test_every_deck_and_environment_setting_has_a_check():
    assert set(FULL_DECK) == set(DECK_KEYS[""])
    for table, names in DECK_KEYS.items():
        if table not in ("", "inputs.documents"):
            assert set(FULL_DECK[table]) == set(names), table
    assert set(FULL_DECK["inputs"]["documents"][0]) == set(DECK_KEYS["inputs.documents"])
    assert set(DECK_LOCAL_KEYS) <= set(DECK_KEYS[""])
    assert check_settings("deck", FULL_DECK, file="deck.toml") == []
    assert check_deck_requirements(FULL_DECK, file="deck.toml") == []
    assert set(FULL_ENVIRONMENT) == set(ENVIRONMENT_KEYS[""])
    assert set(FULL_ENVIRONMENT["fonts"]["Yu Gothic UI"]) == set(ENVIRONMENT_KEYS["fonts.*"])
    for name in ENVIRONMENT_KEYS["renderers"]:
        assert set(FULL_ENVIRONMENT["renderers"][name]) == set(ENVIRONMENT_KEYS[f"renderers.{name}"])
    assert check_settings("environment", FULL_ENVIRONMENT, file="environment.toml") == []


@pytest.mark.parametrize(
    ("data", "expected"),
    [
        ({"entrypont": "slides.py:build"}, "deck.toml: entrypont は不明な設定です（候補: entrypoint）"),
        ({"render": {"requird": ["powerpoint"]}}, "deck.toml: render.requird は不明な設定です（候補: required）"),
        ({"math": {"fallbak": "auto"}}, "deck.toml: math.fallbak は不明な設定です（候補: fallback、fallback_scale）"),
        ({"images": {"raster_scal": 2}}, "deck.toml: images.raster_scal は不明な設定です（候補: raster_scale）"),
        ({"metadata": {"titel": "題名"}}, "deck.toml: metadata.titel は不明な設定です（候補: title）"),
        ({"schema_version": 2}, "deck.toml: schema_version は 1 で指定してください（値: 2）"),
        ({"extends": 3}, "deck.toml: extends は deck.toml のパスで指定してください（値: 3）"),
        (
            {"entrypoint": "slides.py"},
            "deck.toml: entrypoint は slides.py:build の形式で指定してください（値: 'slides.py'）",
        ),
        (
            {"entrypoint": "slides.py:build-all"},
            "deck.toml: entrypoint は slides.py:build の形式で指定してください（値: 'slides.py:build-all'）",
        ),
        # 使われていなかった {width, height} の形は受け付けない。
        (
            {"size": {"width": 960, "height": 540}},
            "deck.toml: size は 16:9、4:3、または [幅, 高さ]（正の数、pt）で指定してください（値: {'width': 960, 'height': 540}）",
        ),
        (
            {"size": [960, 0]},
            "deck.toml: size は 16:9、4:3、または [幅, 高さ]（正の数、pt）で指定してください（値: [960, 0]）",
        ),
        ({"theme": ""}, "deck.toml: theme は設定名またはファイルのパスで指定してください（値: ''）"),
        (
            {"output_dir": "https://example.invalid/out"},
            "deck.toml: output_dir はローカルのパスで指定してください（外部 URL は使えません）（値: 'https://example.invalid/out'）",
        ),
        ({"metadata": []}, "deck.toml: metadata は表形式で指定してください（値: []）"),
        ({"metadata": {"subject": 3}}, "deck.toml: metadata.subject は文字列で指定してください（値: 3）"),
        ({"inputs": {"python": "slides.py"}}, "deck.toml: inputs.python は配列で指定してください（値: 'slides.py'）"),
        (
            {"inputs": {"read_roots": [".", ""]}},
            "deck.toml: inputs.read_roots の 2 番目は空でない文字列で指定してください（値: ''）",
        ),
        (
            {"render": {"primary": "viewer"}},
            "deck.toml: render.primary は powerpoint、libreoffice のどれかで指定してください（値: 'viewer'）",
        ),
        (
            {"render": {"required": []}},
            "deck.toml: render.required には powerpoint、libreoffice のうち 1 つ以上を指定してください",
        ),
        (
            {"render": {"required": ["powerpoint", "powerpoint"]}},
            "deck.toml: render.required の 2 番目の powerpoint は重複しています",
        ),
        (
            {"render": {"required": ["libreofice"]}},
            "deck.toml: render.required の 1 番目は powerpoint、libreoffice のどれかで指定してください（値: 'libreofice'）（候補: libreoffice）",
        ),
        (
            {"math": {"fallback": "always"}},
            "deck.toml: math.fallback は auto、required、off のどれかで指定してください（値: 'always'）",
        ),
        (
            {"math": {"fallback_renderer": "libreoffice"}},
            "deck.toml: math.fallback_renderer は powerpoint で指定してください（互換画像は PowerPoint で作ります）（値: 'libreoffice'）",
        ),
        (
            {"math": {"fallback_scale": 1}},
            "deck.toml: math.fallback_scale は 1 より大きく 8 以下の数で指定してください（値: 1）",
        ),
        (
            {"math": {"functions": ["argmax", "arg max"]}},
            "deck.toml: math.functions の 2 番目は英字だけで指定してください（値: 'arg max'）",
        ),
        (
            {"images": {"raster_scale": 0.5}},
            "deck.toml: images.raster_scale は 1 以上 16 以下の数で指定してください（値: 0.5）",
        ),
        ({"images": {"vector": "yes"}}, "deck.toml: images.vector は真偽値で指定してください（値: 'yes'）"),
        (
            {"template": {"path": "template.pptx", "layout_index": -1}},
            "deck.toml: template.layout_index は 0 以上の整数で指定してください（値: -1）",
        ),
        (
            {"template": {"path": "template.pptx", "show_master_shapes": "no"}},
            "deck.toml: template.show_master_shapes は真偽値で指定してください（値: 'no'）",
        ),
        ({"deliverables": {"pptx": ""}}, "deck.toml: deliverables.pptx は空でない文字列で指定してください（値: ''）"),
    ],
)
def test_deck_setting_names_and_values(data, expected):
    assert [str(p) for p in check_settings("deck", data, file="deck.toml")] == [expected]


def test_deck_documents_are_checked_per_item():
    documents = [
        {"id": "outline", "path": "outline.md"},
        {"id": "outline", "path": "notes.md"},
        {"id": "", "pth": "x.md"},
        "notes.md",
    ]
    assert [str(p) for p in check_settings("deck", {"inputs": {"documents": documents}}, file="deck.toml")] == [
        "deck.toml: inputs.documents の 2 番目の id outline は 1 番目と重複しています",
        "deck.toml: inputs.documents の 3 番目の pth は不明な設定です（候補: path）",
        "deck.toml: inputs.documents の 3 番目に path を指定してください",
        "deck.toml: inputs.documents の 3 番目の id は空でない文字列で指定してください（値: ''）",
        "deck.toml: inputs.documents の 4 番目は表形式で指定してください（値: 'notes.md'）",
    ]


def test_deck_requirements_after_merge():
    assert [str(p) for p in check_deck_requirements({}, file="deck.toml")] == [
        "deck.toml: metadata.title を明示指定してください",
        "deck.toml: metadata.author を明示指定してください",
        "deck.toml: metadata.last_modified_by を明示指定してください",
    ]
    metadata = {"title": " ", "author": "", "last_modified_by": ""}
    assert [str(p) for p in check_deck_requirements({"metadata": metadata}, file="deck.toml")] == [
        "deck.toml: metadata.title を明示指定してください"
    ]
    config = {
        "metadata": {**metadata, "title": "題名"},
        "render": {"primary": "libreoffice"},
        "template": {"layout_index": 6},
    }
    assert [str(p) for p in check_deck_requirements(config, file="deck.toml")] == [
        "deck.toml: render.primary の libreoffice を render.required に含めてください",
        "deck.toml: template.path を指定してください",
    ]
    # extends で合成したときは値を書いたファイルを示し、値がなければその表を書いたファイルか資料自身を示す。
    origins = {
        "render": "../base/deck.toml",
        "render.primary": "../base/deck.toml",
        "metadata": "deck.toml",
        "metadata.title": "deck.toml",
    }
    config = {"metadata": {"title": "題名"}, "render": {"primary": "libreoffice", "required": ["powerpoint"]}}
    assert [str(p) for p in check_deck_requirements(config, file="deck.toml", origins=origins)] == [
        "deck.toml: metadata.author を明示指定してください",
        "deck.toml: metadata.last_modified_by を明示指定してください",
        "../base/deck.toml: render.primary の libreoffice を render.required に含めてください",
    ]
    # 形の間違った値は check_settings が報告するので、ここでは飛ばす。
    malformed = {"metadata": [], "render": {"primary": 3, "required": "powerpoint"}, "template": "template.pptx"}
    assert check_deck_requirements(malformed, file="deck.toml") == []


@pytest.mark.parametrize(
    ("data", "expected"),
    [
        # 使われていなかった設定は受け付けない。
        ({"temporary_dir": "tmp"}, "environment.toml: temporary_dir は不明な設定です"),
        ({"cache_dir": "cache"}, "environment.toml: cache_dir は不明な設定です"),
        (
            {"fonts": {"Arial": {"path": "arial.ttf", "windows_path": "C:\\Windows\\Fonts\\arial.ttf"}}},
            "environment.toml: fonts.Arial.windows_path は不明な設定です",
        ),
        (
            {
                "fonts": {
                    "Yu Gothic UI": {"path": "YuGothM.ttc", "windows_bold_path": "C:\\Windows\\Fonts\\YuGothB.ttc"}
                }
            },
            'environment.toml: fonts."Yu Gothic UI".windows_bold_path は不明な設定です（候補: bold_path）',
        ),
        ({"jbos": 2}, "environment.toml: jbos は不明な設定です（候補: jobs）"),
        ({"renderers": {"powerpnt": {}}}, "environment.toml: renderers.powerpnt は不明な設定です（候補: powerpoint）"),
        (
            {"renderers": {"powerpoint": {"timout": 30}}},
            "environment.toml: renderers.powerpoint.timout は不明な設定です（候補: timeout）",
        ),
        (
            {"renderers": {"libreoffice": {"wait_seconds": 5}}},
            "environment.toml: renderers.libreoffice.wait_seconds は不明な設定です",
        ),
        ({"schema_version": 2}, "environment.toml: schema_version は 1 で指定してください（値: 2）"),
        ({"jobs": 1.5}, "environment.toml: jobs は 1 以上の整数で指定してください（値: 1.5）"),
        (
            {"font_directories": "fonts"},
            "environment.toml: font_directories は文字列の配列で指定してください（値: 'fonts'）",
        ),
        (
            {"font_directories": ["https://example.invalid/fonts"]},
            "environment.toml: font_directories の 1 番目はローカルのパスで指定してください（外部 URL は使えません）"
            "（値: 'https://example.invalid/fonts'）",
        ),
        (
            {"fonts": {"Arial": "arial.ttf"}},
            "environment.toml: fonts.Arial は表形式で指定してください（値: 'arial.ttf'）",
        ),
        (
            {"fonts": {"Arial": {"path": ""}}},
            "environment.toml: fonts.Arial.path は空でない文字列で指定してください（値: ''）",
        ),
        (
            {"fonts": {"Arial": {"path": "arial.ttf", "index": -1}}},
            "environment.toml: fonts.Arial.index は 0 以上の整数で指定してください（値: -1）",
        ),
        (
            {"fonts": {"Arial": {"path": "arial.ttf", "bold_index": True}}},
            "environment.toml: fonts.Arial.bold_index は 0 以上の整数で指定してください（値: True）",
        ),
        (
            {"renderers": {"powerpoint": {"timeout": 0}}},
            "environment.toml: renderers.powerpoint.timeout は正の秒数で指定してください（値: 0）",
        ),
        (
            {"renderers": {"powerpoint": {"wait_seconds": -1}}},
            "environment.toml: renderers.powerpoint.wait_seconds は 0 以上の秒数で指定してください（値: -1）",
        ),
        (
            {"renderers": {"libreoffice": {"executable": ""}}},
            "environment.toml: renderers.libreoffice.executable は空でない文字列で指定してください（値: ''）",
        ),
    ],
)
def test_environment_setting_names_and_values(data, expected):
    assert [str(p) for p in check_settings("environment", data, file="environment.toml")] == [expected]


VALID_DECORATIONS = [
    {
        "kind": "line",
        "id": "rule",
        "start": [0.10, 0.67],
        "end": [0.90, 0.67],
        "line_color": "primary",
        "line_width_pt": 3,
    },
    {
        "kind": "box",
        "id": "mark",
        "shape": "diamond",
        "area_pt": [86, 356, 10, 10],
        "fill": "primary",
        "line_color": "none",
    },
    {
        "kind": "box",
        "id": "band",
        "shape": "rounded_rectangle",
        "area": [0, 0.9, 1, 0.1],
        "corner_radius_pt": 4,
        "dashed": True,
    },
    {"kind": "arrow", "id": "pointer", "start_pt": [10, 20], "end_pt": [60, 20]},
]


def test_valid_decorations_pass():
    assert check_decorations(VALID_DECORATIONS, page_kind="section", file="theme.toml") == []
    theme = {"page_kinds": {"section": {"decorations": VALID_DECORATIONS}}}
    assert problems_of("theme", theme) == []


@pytest.mark.parametrize(
    ("item", "expected"),
    [
        ({"id": "rule", "start": [0, 0], "end": [1, 1]}, "の 2 番目に kind（line、arrow、box）を指定してください"),
        (
            {"kind": "circle", "id": "c"},
            "の 2 番目の kind は line、arrow、box のどれかで指定してください（値: 'circle'）",
        ),
        ({"kind": "line", "start": [0, 0], "end": [1, 1]}, "の 2 番目に id を指定してください"),
        (
            {"kind": "line", "id": "rule", "start": [0, 0], "end": [1, 1]},
            "の 2 番目の id rule は 1 番目と重複しています",
        ),
        (
            {"kind": "line", "id": "a/b", "start": [0, 0], "end": [1, 1]},
            "の 2 番目の id は空でなく、/ と制御文字を含まない文字列にしてください（値: 'a/b'）",
        ),
        (
            {"kind": "box", "id": "b", "area": [0, 0, 1, 1], "area_pt": [0, 0, 10, 10]},
            "の 2 番目には area か area_pt のどちらか一方を指定してください",
        ),
        ({"kind": "box", "id": "b"}, "の 2 番目には area か area_pt のどちらか一方を指定してください"),
        (
            {"kind": "line", "id": "l", "start": [0, 0], "end_pt": [10, 10]},
            "の 2 番目には start と end か、start_pt と end_pt のどちらか一方の組を指定してください",
        ),
        ({"kind": "line", "id": "l", "start": [0, 0]}, "の 2 番目には start と end の両方を指定してください"),
        (
            {"kind": "line", "id": "l", "start": [0, 0], "end": [1.5, 0]},
            "の 2 番目の end は 0 以上 1 以下の比で指定してください（値: [1.5, 0]）",
        ),
        ({"kind": "arrow", "id": "l", "start_pt": [5, 5], "end_pt": [5, 5]}, "の 2 番目の始点と終点を別にしてください"),
        (
            {"kind": "box", "id": "b", "area": [0, 0, 0.5], "text": "見出し"},
            "の 2 番目の text は使えません。飾りには文字を置きません",
        ),
        (
            {"kind": "line", "id": "l", "start": [0, 0], "end": [1, 1], "fill": "primary"},
            "の 2 番目の fill は line の飾りでは使えません",
        ),
        (
            {"kind": "line", "id": "l", "start": [0, 0], "end": [1, 1], "area": [0, 0, 1, 1]},
            "の 2 番目の area は line の飾りでは使えません",
        ),
        (
            {"kind": "box", "id": "b", "area": [0, 0, 1, 1], "corner_radius_pt": 4},
            "の 2 番目の corner_radius_pt は shape が rounded_rectangle のときだけ指定できます",
        ),
        (
            {"kind": "box", "id": "b", "area": [0, 0, 1, 1], "shape": "star"},
            "の 2 番目の shape は rectangle、rounded_rectangle、ellipse、diamond のどれかで指定してください（値: 'star'）",
        ),
        (
            {"kind": "box", "id": "b", "area": [0, 0, 1, 1], "fill_colour": "primary"},
            "の 2 番目の fill_colour は不明な設定です（候補: fill）",
        ),
        (
            {"kind": "line", "id": "l", "start": [0, 0], "end": [1, 1], "line_width_pt": -1},
            "の 2 番目の line_width_pt は 0 以上の有限の数で指定してください（値: -1）",
        ),
    ],
)
def test_invalid_decorations_are_rejected(item, expected):
    decorations = [VALID_DECORATIONS[0], item]
    problems = [str(p) for p in check_decorations(decorations, page_kind="section", file="theme.toml")]
    assert "theme.toml: page_kinds.section.decorations " + expected in problems


def test_decoration_problems_come_through_theme_checks():
    assert problems_of("theme", {"page_kinds": {"section": {"decorations": {"kind": "line"}}}}) == [
        "theme.toml: page_kinds.section.decorations は配列で指定してください（値: {'kind': 'line'}）"
    ]
    data = {"page_kinds": {"section": {"decorations": [{**VALID_DECORATIONS[0], "line_color": "primery"}]}}}
    assert problems_of("theme", data) == [
        "theme.toml: page_kinds.section.decorations の 1 番目の line_color の色 primery はテーマにありません（候補: primary）"
    ]
    # 線に書いた fill は対応していない項目として 1 回だけ報告し、色の参照としては報告しない。種類の間違った飾りの色も調べない。
    data = {
        "page_kinds": {
            "section": {
                "decorations": [
                    {**VALID_DECORATIONS[0], "fill": "unknown"},
                    {"kind": ["box"], "id": "x", "fill": "unknown"},
                ]
            }
        }
    }
    assert problems_of("theme", data) == [
        "theme.toml: page_kinds.section.decorations の 1 番目の fill は line の飾りでは使えません",
        "theme.toml: page_kinds.section.decorations の 2 番目の kind は line、arrow、box のどれかで指定してください（値: ['box']）",
    ]


def text(value):
    return {"kind": "text", "text": value}


def math(value):
    return {"kind": "math", "text": value}


def test_role_usage_problems_detect_restricted_math_only_and_text_only():
    listed = {"size_pt": 18, "min_size_pt": 16, "allowed_slide_ids": ["results"]}
    assert role_usage_problems("small", listed, slide_id="results", declared=(), content=[text("説明")]) == []
    assert role_usage_problems("small", listed, slide_id="intro", declared=(), content=[text("説明")]) == [
        RoleProblem("restricted", "small")
    ]
    assert role_usage_problems("small", listed, slide_id="intro", declared={"small"}, content=[]) == []
    restricted = {"size_pt": 18, "min_size_pt": 16, "restricted": True}
    assert role_usage_problems("small", restricted, slide_id="results", declared=("other",), content=[]) == [
        RoleProblem("restricted", "small")
    ]
    assert (
        role_usage_problems(
            "small", {**restricted, "allowed_slide_ids": ["results"]}, slide_id="results", declared=(), content=[]
        )
        == []
    )
    assert (
        role_usage_problems(
            "body", {"size_pt": 20, "min_size_pt": 18, "restricted": False}, slide_id="intro", declared=(), content=[]
        )
        == []
    )
    formula = {"size_pt": 30, "min_size_pt": 24, "allowed_content": "math"}
    assert role_usage_problems("formula", formula, slide_id="intro", declared=(), content=[math("x"), text(" ")]) == []
    assert role_usage_problems(
        "formula", formula, slide_id="intro", declared=(), content=[math("x"), text("と y")]
    ) == [RoleProblem("math_only", "formula")]
    caption = {"size_pt": 14, "min_size_pt": 12, "allowed_content": "text"}
    assert role_usage_problems("caption", caption, slide_id="intro", declared=(), content=[text("図 1")]) == []
    assert role_usage_problems("caption", caption, slide_id="intro", declared=(), content=[text("図 "), math("n")]) == [
        RoleProblem("text_only", "caption")
    ]
    both = {**formula, "allowed_slide_ids": ["results"]}
    assert [
        p.kind for p in role_usage_problems("formula", both, slide_id="intro", declared=(), content=[text("説明")])
    ] == ["restricted", "math_only"]
    # 配列でない allowed_slide_ids（設定の誤り）は、部分一致でページを許さない。
    assert role_usage_problems(
        "small", {**listed, "allowed_slide_ids": "results"}, slide_id="result", declared=(), content=[]
    ) == [RoleProblem("restricted", "small")]


def test_role_problem_messages():
    assert RoleProblem("restricted", "formula").describe(setting="theme.toml の roles.formula") == (
        '役割 formula はこのページでは使えません。deck.slide(..., allow_roles=["formula"]) か要素の allow_roles で宣言してください'
        "（設定: theme.toml の roles.formula）"
    )
    assert RoleProblem("math_only", "formula").describe() == "役割 formula は数式だけに使えます"
    assert RoleProblem("text_only", "caption").describe() == "役割 caption に数式は使えません"


def test_page_kinds_are_the_supported_kinds():
    assert set(PAGE_KINDS) == set(BUNDLED_THEME["page_kinds"])


def test_checks_do_not_change_their_input():
    data = {
        "roles": {"note": {"size": 14, "color": "mutd"}},
        "page_kinds": {"section": {"decorations": copy.deepcopy(VALID_DECORATIONS)}},
    }
    before = copy.deepcopy(data)
    problems_of("theme", data)
    assert data == before


def test_role_skip_checks_accept_only_warning_and_info_codes():
    from slidekit.schema import check_settings

    theme = {"roles": {"refs": {"skip_checks": ["math_not_marked", "text_overflow", "no_such_code"]}}}
    messages = [str(problem) for problem in check_settings("theme", theme, file="theme.toml")]
    assert any(
        "roles.refs.skip_checks" in message and "text_overflow は指定できません" in message for message in messages
    )
    assert any("診断の識別子で指定してください" in message for message in messages)
    assert not any("math_not_marked" in message for message in messages)
