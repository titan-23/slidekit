"""資料の設定と、公開 API が参照したローカル入力を管理する。

`load_project` は 2 段階で読む。先に設定ファイル（`deck.toml` と `extends` の継承元、テーマ、記述規則、実行環境）の
設定名と値を確かめ、問題をすべて集めて 1 つの `schema.SettingsError` にする。継承をたどれなければ（`extends` の誤り、
継承元の構文の誤り）、読み込みに成功した `deck.toml` の問題と合わせて、そこで報告する。合成した後の検査（参照、必須の設定、
書体の path）は、関わる設定ファイルの読み込みがすべて成功したときだけ行う（読み込みに失敗した原因を、間違った問題で隠さないため）。
その後でパスを解決し、参照範囲・ファイルの有無・出力先との重なりを確かめて入力を追跡する（この段階の誤りは、見つけた時点で例外にする）。
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from functools import cached_property
import os
from pathlib import Path
import re
import tomllib

from . import schema
from .files import file_hash
from .layout import finite_number
from .errors import InputError
from .schema import Problem, SettingsError

PACKAGE = Path(__file__).resolve().parent
# 同梱の設定。種類から (フォルダー, 設定名)。
BUNDLED = {"theme": ("themes", "default"), "rules": ("rules", "general")}
# extends をたどる回数の上限。
MAX_EXTENDS = 8
_URL = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*://")


def merge(base, override):
    result = copy.deepcopy(base)
    for key, value in override.items():
        result[key] = (
            merge(result[key], value)
            if isinstance(value, dict) and isinstance(result.get(key), dict)
            else copy.deepcopy(value)
        )
    return result


def _read(path, name=None):
    """TOML の設定ファイルを読む。構文の誤りはそのファイル（表示名 name）の問題として SettingsError にする。"""
    try:
        with Path(path).open("rb") as stream:
            return tomllib.load(stream)
    except (tomllib.TOMLDecodeError, UnicodeDecodeError) as error:
        raise SettingsError([Problem(name or str(path), "", f"TOML として読めません（{error}）")]) from error


def _within(path, root):
    return path == root or root in path.parents


def _display(path, root=None):
    """問題を示すときのファイル名。同梱の設定は「同梱の themes/default.toml」、資料のディレクトリの中は相対パス、ほかは絶対パス。"""
    path = Path(path)
    if _within(path, PACKAGE):
        return f"同梱の {path.relative_to(PACKAGE).as_posix()}"
    if root is not None and _within(path, Path(root)):
        return path.relative_to(root).as_posix()
    return str(path)


def _is_local(value):
    return isinstance(value, str) and bool(value.strip()) and not _URL.match(value)


def _local_path(value, base):
    if not isinstance(value, (str, Path)) or not str(value).strip():
        raise ValueError("ローカルのパスを指定してください")
    if _URL.match(str(value)):
        raise ValueError("外部 URL は入力ファイルとして使えません")
    value = Path(value).expanduser()
    return (value if value.is_absolute() else base / value).resolve()


def _check_resource(project, path):
    """ライブラリと実行環境の入力が、出力先と重ならないことを確かめる。"""
    path = Path(path).resolve()
    if str(path) in project.config["deliverables"].values() or _within(path, project.output_dir):
        raise ValueError(f"ライブラリ・実行環境の入力と出力先が重複しています: {path}")
    return path


def _track_resource(project, path):
    """ライブラリと実行環境は、資料用の参照範囲とは別に追跡する。"""
    path = _check_resource(project, path)
    if path.is_file():
        project.input_files[str(path)] = file_hash(path)


@dataclass
class Project:
    path: Path
    root: Path
    config: dict
    theme: dict
    rules: dict
    environment: dict
    input_files: dict = field(default_factory=dict)
    output_dir: Path = Path("out")
    # 資料の説明（brief.md）。資料のディレクトリから継承元の順に探し、なければ None。
    brief_path: Path | None = None
    # 続行できる入力の誤りを、例外にせず場面の issues に集めるか（構築では真にする）。
    collect_errors: bool = False
    # 部分の生成・検査で選んだページ（--slides の指定の文字列）。指定がなければ None。生成の流れが構築の前に入れる。
    selection: str | None = None

    def selected(self, *slide_ids) -> bool:
        """ページの識別子のどれかが選択（--slides）に含まれるか。選択がなければ常に真を返す。

        資料のコードは、選択があるときに、選ばれたページを含む章だけを作ってよい（ほかの章の書きかけの誤りで止まらない）。
        選択にページ番号（3、5-7）があれば、番号は全ページを作らないと決まらないので、常に真を返す。
        偽を返した識別子のページを作らなかったときは、生成の流れが「選んだ章だけを構築した」とみなし、
        ページを番号でなく識別子で示す（作ったページの中の番号は、資料の元のページ番号と違うため）。
        """
        if not slide_ids:
            raise TypeError("selected() にはページの識別子を 1 つ以上渡してください")
        if self.selection is None:
            return True
        cached = self.__dict__.get("_selection_items")
        if cached is None or cached[0] != self.selection:
            from .selection import parse

            cached = self._selection_items = (self.selection, parse(self.selection))
        items = cached[1]
        if any(item[0] == "range" for item in items):
            return True
        chosen = {item[1] for item in items}
        if any(slide_id in chosen for slide_id in slide_ids):
            return True
        # 構築の後に、選ばなかったページを作らなかったかを調べるために残す（pipeline が読む）。
        self.__dict__.setdefault("_unselected", set()).update(slide_ids)
        return False

    @property
    def size(self):
        return tuple(self.config["size"])

    @property
    def read_roots(self):
        return tuple(Path(p) for p in self.config["inputs"]["read_roots"])

    def resolve_input(self, value) -> Path:
        path = _local_path(value, self.root)
        if not any(_within(path, root) for root in self.read_roots):
            raise ValueError(f"宣言した参照範囲外です: {path}")
        if not path.is_file():
            raise FileNotFoundError(f"入力ファイルが見つかりません: {path}")
        destinations = (Path(p) for p in self.config.get("deliverables", {}).values())
        if any(path == destination for destination in destinations) or _within(path, self.output_dir):
            raise ValueError(f"出力先を入力として使うことはできません: {path}")
        self.input_files[str(path)] = file_hash(path)
        return path

    @property
    def fonts(self):
        """寸法の計算に使う書体（measure.FontSource）。初回の参照で FontRegistry を作り、構築・検査・書き出しで共有する。

        試験では代入で偽物に差し替えることができる。閉じる処理は、使う側（生成の流れ）が最後に 1 回行う。
        """
        registry = self.__dict__.get("_font_registry")
        if registry is None:
            from .fonts import FontRegistry

            registry = self._font_registry = FontRegistry(self)
        return registry

    @fonts.setter
    def fonts(self, value):
        self._font_registry = value

    @cached_property
    def bibliography(self):
        """inputs.bibliography の文献（bibliography.Bibliography）。初回の参照で読み、指定がなければ空の Bibliography を返す。"""
        from .bibliography import load_bibliography

        return load_bibliography(self.config["inputs"]["bibliography"])

    def measure_text(self, text, size, bold=False, **options):
        """互換用。(最大の行幅, 高さ) を self.fonts で求める。"""
        return self.fonts.measure_text(text, size, bold, **options)

    def measure_math(self, text, size, bold=False, **options):
        """互換用。数式の (幅, 高さ) を self.fonts で求める。"""
        return self.fonts.measure_math(text, size, bold, **options)


def _size(value):
    """ページの寸法（pt）。値の形は読み込みの最初に確かめてある。"""
    if isinstance(value, str):
        return [960.0, 540.0] if value == "16:9" else [720.0, 540.0]
    return [finite_number(number, "画面寸法") for number in value]


def _resolve_roles(theme, file):
    """役割の継承を解決し、文字サイズを数値にそろえる。設定の検査を通ったテーマでは問題は出ない。問題の一覧を返す。"""
    roles, resolved, problems = theme["roles"], {}, []

    def resolve(name, active):
        if name in resolved:
            return resolved[name]
        role = copy.deepcopy(roles[name])
        parent = role.pop("inherits", None)
        if parent:
            if parent not in roles or parent in active | {name}:
                problems.append(Problem(file, f"roles.{name}.inherits", f"の継承が不正です（{parent}）"))
                return None
            base = resolve(parent, active | {name})
            if base is None:
                return None
            role = merge(base, role)
        try:
            for key in ("size_pt", "min_size_pt"):
                role[key] = finite_number(role[key], f"{name}.{key}")
        except (KeyError, ValueError):
            problems.append(Problem(file, f"roles.{name}", "に size_pt と min_size_pt を数で指定してください"))
            return None
        resolved[name] = role
        return role

    for name in roles:
        resolve(name, set())
    theme["roles"] = resolved
    return problems


def _environment_paths(data, base, name):
    """実行環境の設定の相対パスを、その設定ファイル（表示名 name）の場所を基準に絶対パスへ直す。

    形の間違った値はそのまま残す（設定の検査が報告する）。戻り値は (設定, 問題の一覧) で、問題は font_directories のないフォルダー。
    """
    data, problems = copy.deepcopy(data), []
    fonts = data.get("fonts")
    for font in fonts.values() if isinstance(fonts, dict) else ():
        if isinstance(font, dict):
            for key in ("path", "bold_path"):
                if _is_local(font.get(key)):
                    font[key] = str(_local_path(font[key], base))
    folders = data.get("font_directories")
    if isinstance(folders, list) and all(_is_local(folder) for folder in folders):
        data["font_directories"] = []
        for folder in folders:
            folder = _local_path(folder, base)
            if not folder.is_dir():
                problems.append(Problem(name, "font_directories", f"のフォルダーがありません: {folder}"))
            data["font_directories"].append(str(folder))
    powerpoint = data.get("powerpoint")
    value = powerpoint.get("powershell") if isinstance(powerpoint, dict) else None
    if _is_local(value) and ("/" in value or "\\" in value) and not re.match(r"^[A-Za-z]:", value):
        powerpoint["powershell"] = str(_local_path(value, base))
    return data, problems


def _common_environment_path():
    """共通の実行環境の設定のパス。SLIDEKIT_ENVIRONMENT か利用者ごとの既定の場所にあるものを使い、なければ None。"""
    default = (
        Path(os.environ.get("APPDATA", str(Path.home() / "AppData/Roaming"))) / "slidekit/environment.toml"
        if os.name == "nt"
        else Path.home() / ".config/slidekit/environment.toml"
    )
    specified = os.environ.get("SLIDEKIT_ENVIRONMENT")
    common = _local_path(specified, Path.cwd()) if specified else default
    if specified and not common.is_file():
        raise FileNotFoundError(f"実行環境の設定が見つかりません: {common}")
    return common if common.is_file() else None


def _environment_settings(paths, root=None, *, complete=True):
    """既定値に実行環境の設定ファイルを順に重ねる。戻り値は (設定, 問題の一覧)。読んだファイルを config_files に記録する。

    ファイルごとに設定名と値を確かめる。重ねた後に書体に path があることは、すべての設定ファイルの読み込みに成功したときだけ
    確かめる（complete が偽なら、呼び出し側で読み込みに失敗した設定ファイルがある）。
    """
    data = {"fonts": {}, "powerpoint": {"powershell": "powershell.exe", "timeout": 90}}
    records, problems, owners = [], [], {}
    for path in paths:
        name = _display(path, root)
        try:
            own = _read(path, name)
        except SettingsError as error:
            problems.extend(error.problems)
            complete = False
            continue
        problems.extend(schema.check_settings("environment", own, file=name))
        own, found = _environment_paths(own, path.parent, name)
        problems.extend(found)
        data = merge(data, own)
        owners.update(dict.fromkeys(own["fonts"] if isinstance(own.get("fonts"), dict) else (), name))
        records.append({"path": str(path), "sha256": file_hash(path)})
    fonts = data["fonts"] if complete and isinstance(data.get("fonts"), dict) else {}
    for family, font in fonts.items():
        if isinstance(font, dict) and "path" not in font:
            problems.append(
                Problem(owners.get(family, ""), f"fonts.{schema.toml_key(family)}", "に path を指定してください")
            )
    data["config_files"] = records
    return data, problems


def _environment_data(paths):
    """既定値に設定ファイルを順に重ね、読んだファイルを config_files に記録する。設定の誤りは 1 つの SettingsError にまとめる。"""
    data, problems = _environment_settings(paths)
    if problems:
        raise SettingsError(problems)
    return data


def load_common_environment() -> dict:
    """資料に依らない実行環境（既定値と共通の environment.toml）を返す。書体の探索と入力の追跡は行わない。"""
    path = _common_environment_path()
    return _environment_data([path] if path else [])


_FONT_KEYS = ("font_japanese", "font_latin", "font_math")


def _theme_families(theme):
    """テーマが使う書体名を、現れた順に返す（fonts と、書式の font_japanese・font_latin・font_math）。"""
    families = [theme["fonts"][name] for name in ("japanese", "latin", "math")]

    def collect(value):
        if isinstance(value, dict):
            for key, item in value.items():
                if key in _FONT_KEYS and isinstance(item, str) and item:
                    families.append(item)
                else:
                    collect(item)
        elif isinstance(value, list):
            for item in value:
                collect(item)

    collect({key: value for key, value in theme.items() if key != "fonts"})
    return list(dict.fromkeys(families))


def _load_environment(project, data, paths):
    """資料の実行環境。設定ファイルを入力として追跡し、テーマが使う書体のうち指定のないものを既定の候補から探す。"""
    from .fonts import default_font

    for path in paths:
        _track_resource(project, path)
    # テーマが使う書体のうち、指定のないものだけを既定の候補から探す。使わない書体は探さない。
    for family in _theme_families(project.theme):
        if family not in data["fonts"]:
            specification = default_font(family, data)
            if specification is not None:
                data["fonts"][family] = specification
    # 書体のハッシュ値はここでは求めない。計測に使った書体だけを、生成の記録が追跡する。
    for font in data["fonts"].values():
        for key in ("path", "bold_path"):
            if font.get(key):
                _check_resource(project, font[key])
    return data


def _template(project):
    """雛形の PPTX が空白レイアウトを持つことを確かめ、template["inspection"] に記録する。設定の値は読み込みの最初に確かめてある。"""
    template = project.config.get("template")
    if not template:
        return
    path = project.resolve_input(template["path"])
    if path.suffix.lower() != ".pptx":
        raise ValueError("雛形は PPTX ファイルを指定してください")
    template["path"] = str(path)
    index = template.setdefault("layout_index", 6)
    template.setdefault("show_master_shapes", False)
    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER

    presentation = Presentation(path)
    if index >= len(presentation.slide_layouts):
        raise ValueError("雛形のレイアウト番号が範囲外です")
    allowed = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}
    if any(
        not shape.is_placeholder or shape.placeholder_format.type not in allowed
        for shape in presentation.slide_layouts[index].shapes
    ):
        raise ValueError("雛形には空白レイアウトを指定してください")
    template["inspection"] = {
        "layout_index": index,
        "layout_name": presentation.slide_layouts[index].name,
        "original_slide_count": len(presentation.slides),
        "blank_layout": True,
    }


def _read_chain(path):
    """deck.toml から extends をたどって読み、ファイルごとに設定名と値を確かめる。戻り値は (設定の列, 問題の一覧, たどり終えたか)。

    設定の列は [(パス, 設定), …] で、資料自身から継承元の順に並ぶ。extends は deck.toml のファイルかそのディレクトリで、
    相対パスは extends を書いたファイルの場所を基準にする。たどるのは MAX_EXTENDS 回まで。構文の誤り・extends の形の誤り・
    循環・回数の超過・ファイルがないときは、そこでたどるのをやめ、読み込みに成功したファイルの問題に加えて返す。
    """
    root = path.parent
    chain, problems, current = [], [], path
    while True:
        name = _display(current, root)
        try:
            data = _read(current, name)
        except SettingsError as error:
            return chain, problems + list(error.problems), False
        chain.append((current, data))
        problems += schema.check_settings("deck", data, file=name)
        if "extends" not in data:
            return chain, problems, True
        if not _is_local(data["extends"]):
            # 形の誤りは、上の設定の検査が報告している。
            return chain, problems, False
        parent = _local_path(data["extends"], current.parent)
        if parent.is_dir():
            parent = (parent / "deck.toml").resolve()
        if any(parent == file for file, _ in chain):
            message = "が循環しています"
        elif len(chain) > MAX_EXTENDS:
            message = f"の継承は {MAX_EXTENDS} 段までにしてください"
        elif not parent.is_file():
            message = "の deck.toml がありません"
        else:
            current = parent
            continue
        return chain, problems + [Problem(name, "extends", f"{message}: {parent}")], False


def _absolute_paths(data, base):
    """継承元の deck.toml に書いたパスを、そのファイルのディレクトリ base を基準に絶対パスへ直す。形の間違った値はそのまま残す。

    対象は entrypoint のファイルの部分、theme・rules（同梱の設定名でないとき）、inputs.read_roots・python・bibliography・
    asset_dirs、inputs.documents の path、template.path、environment_override。
    """
    data = copy.deepcopy(data)

    def absolute(value):
        return str(_local_path(value, base)) if _is_local(value) else value

    entrypoint = data.get("entrypoint")
    if isinstance(entrypoint, str) and ":" in entrypoint:
        file, function = entrypoint.rsplit(":", 1)
        data["entrypoint"] = f"{absolute(file)}:{function}"
    for kind, (_, default) in BUNDLED.items():
        if kind in data and data[kind] != default:
            data[kind] = absolute(data[kind])
    if "environment_override" in data:
        data["environment_override"] = absolute(data["environment_override"])
    inputs = data.get("inputs")
    if isinstance(inputs, dict):
        for key in ("read_roots", "python", "bibliography", "asset_dirs"):
            if isinstance(inputs.get(key), list):
                inputs[key] = [absolute(value) for value in inputs[key]]
        for document in inputs.get("documents") if isinstance(inputs.get("documents"), list) else ():
            if isinstance(document, dict) and "path" in document:
                document["path"] = absolute(document["path"])
    template = data.get("template")
    if isinstance(template, dict) and "path" in template:
        template["path"] = absolute(template["path"])
    return data


def _record_origins(data, name, prefix, origins):
    for key, value in data.items():
        position = f"{prefix}.{key}" if prefix else str(key)
        origins[position] = name
        if isinstance(value, dict):
            _record_origins(value, name, position, origins)


def _merge_chain(chain, root):
    """継承元から順に deck.toml を重ねる。表は再帰的に合成し、配列と値は後のファイルが置き換える。

    継承元の output_dir・deliverables は引き継がない。継承元に書いたパスは、そのファイルの場所を基準にした絶対パスにする
    （資料自身のパスは、後で資料のディレクトリを基準に解決する）。戻り値は (設定, 設定の位置から値を書いたファイルの表示名への対応)。
    """
    config, origins = {}, {}
    for index in range(len(chain) - 1, -1, -1):
        file, data = chain[index]
        own = {key: value for key, value in data.items() if key != "extends"}
        if index:
            own = _absolute_paths(
                {key: value for key, value in own.items() if key not in schema.DECK_LOCAL_KEYS}, file.parent
            )
        config = merge(config, own)
        _record_origins(own, _display(file, root), "", origins)
    return config, origins


def _read_roots(config, root, files):
    """参照範囲（解決したパスの文字列の列）。設定の形が間違っていれば None。

    最も近いファイルが指定した列（指定がなければ資料のディレクトリ）に、継承があれば、継承をたどった全ファイルの
    ディレクトリを加える（継承元の入口や補助モジュールの読み込みを許可するため）。
    """
    inputs = config.get("inputs", {})
    values = inputs.get("read_roots", ["."]) if isinstance(inputs, dict) else None
    if not isinstance(values, list) or not all(_is_local(value) for value in values):
        return None
    roots = [str(_local_path(value, root)) for value in values]
    if len(files) > 1:
        for file in files:
            if str(file.parent) not in roots:
                roots.append(str(file.parent))
    return roots


def _readable(value, root, roots):
    """設定に書いた資料のファイルについて、(パス, 参照範囲の中にあって存在するか) を返す。形の間違った値は (None, False)。"""
    if not _is_local(value):
        return None, False
    path = _local_path(value, root)
    inside = roots is not None and any(_within(path, Path(base)) for base in roots)
    return path, inside and path.is_file()


def _bundled(kind):
    folder, default = BUNDLED[kind]
    return PACKAGE / folder / f"{default}.toml"


def _settings(kind, value, root, roots, problems):
    """同梱のテーマか記述規則に、資料の設定ファイルを重ねる。戻り値は (設定, 資料の設定ファイルのパスか None, 読み込みに成功したか)。

    資料のファイルの構文・設定名・値の問題を problems に加える。値の形の誤り、参照範囲の外のファイル、ないファイル、
    構文の誤りでは、同梱の設定だけを返し、3 番目を偽にする（参照範囲とファイルの有無の誤りは、後で resolve_input が報告する）。
    """
    bundled = _bundled(kind)
    data = _read(bundled, _display(bundled))
    if value is None or value == BUNDLED[kind][1]:
        return data, None, True
    source, readable = _readable(value, root, roots)
    if not readable:
        return data, source, False
    name = _display(source, root)
    try:
        own = _read(source, name)
    except SettingsError as error:
        problems.extend(error.problems)
        return data, source, False
    problems.extend(schema.check_settings(kind, own, file=name))
    return merge(data, own), source, True


def load_project(path) -> Project:
    path = Path(path).expanduser().resolve()
    if path.is_dir():
        path /= "deck.toml"
    root = path.parent

    # 1. 設定ファイルの設定名と値。問題をすべて集めてから、1 つの SettingsError にする。
    chain, problems, complete = _read_chain(path)
    if not complete:
        # 継承をたどれなければ、合成した設定を確かめることができない。読み込みに成功した deck.toml の問題と合わせて報告する。
        raise SettingsError(problems)
    files = [file for file, _ in chain]
    config, origins = _merge_chain(chain, root)
    problems += schema.check_deck_requirements(config, file=_display(path, root), origins=origins)
    roots = _read_roots(config, root, files)
    # 参照は、読み込みに成功した設定だけで確かめる。資料のテーマの読み込みに失敗したときに同梱のテーマで確かめると、
    # 失敗の原因の代わりに間違った問題（記述規則が参照する役割がない、など）を示すためである。失敗の原因は、構文の問題か
    # 後の resolve_input が示す。
    theme, theme_source, theme_loaded = _settings("theme", config.get("theme"), root, roots, problems)
    theme_file = _display(theme_source or _bundled("theme"), root)
    if theme_loaded:
        problems += schema.check_theme_references(theme, file=theme_file)
    rules, rules_source, rules_loaded = _settings("rules", config.get("rules"), root, roots, problems)
    if theme_loaded and rules_loaded:
        problems += schema.check_rules_references(rules, theme, file=_display(rules_source or _bundled("rules"), root))
    common = _common_environment_path()
    environment_paths = [common] if common else []
    override, readable = _readable(config.get("environment_override"), root, roots)
    if readable:
        environment_paths.append(override)
    # 上書きの設定を読めなければ、重ねた後の書体を確かめない（同じ理由）。
    environment, found = _environment_settings(
        environment_paths, root, complete=readable or "environment_override" not in config
    )
    problems += found
    if not problems:
        problems += _resolve_roles(theme, theme_file)
    if problems:
        raise SettingsError(problems)

    # 2 の誤り（入力ファイルがない、参照範囲外など）は、設定の場所を示す設定の誤りにする（トレースバックを出さない）。
    where = {"key": ""}

    def resolve(value, key):
        where["key"] = key
        result = project.resolve_input(value)
        where["key"] = ""
        return result

    try:
        # 2. パスの解決、参照範囲・出力先の確認と入力の追跡。
        config.setdefault("schema_version", 1)
        config.setdefault("entrypoint", "slides.py:build")
        config["size"] = _size(config.get("size", "16:9"))
        config["metadata"] = config.get("metadata", {})
        entrypoint = config["entrypoint"]
        inputs = merge(
            {"read_roots": ["."], "python": [], "bibliography": [], "asset_dirs": [], "documents": []},
            config.get("inputs", {}),
        )
        inputs["read_roots"] = roots
        if not inputs["read_roots"] or any(not Path(p).is_dir() for p in inputs["read_roots"]):
            raise ValueError("参照範囲には既存のディレクトリを指定してください")
        config["inputs"] = inputs
        output_dir = _local_path(config.get("output_dir", "out"), root)
        config["output_dir"] = str(output_dir)
        if output_dir.is_file() or output_dir == root or _within(root, output_dir):
            raise ValueError("資料ディレクトリまたはその親を出力先にはできません")
        deliverables = merge(
            {
                "pptx": "presentation.pptx",
                "pdf": "presentation.pdf",
                "notes": "speaker_notes.md",
                "record": "deliverables.json",
            },
            config.get("deliverables", {}),
        )
        config["deliverables"] = {name: str(_local_path(value, root)) for name, value in deliverables.items()}
        destinations = list(map(Path, config["deliverables"].values()))
        if len(set(destinations)) != len(destinations) or any(
            _within(p, output_dir) or p.is_dir() for p in destinations
        ):
            raise ValueError("完成物の保存先同士、実行ディレクトリとの重複は指定できません")
        if any(_within(a, b) or _within(b, a) for i, a in enumerate(destinations) for b in destinations[i + 1 :]):
            raise ValueError("完成物の保存先に親子のパスを指定できません")
        project = Project(path, root, config, theme, rules, {}, output_dir=output_dir)
        for file in files:
            if file in destinations or _within(file, output_dir):
                raise ValueError("資料設定と保存先が重複しています")
            project.input_files[str(file)] = file_hash(file)
        config["math"] = merge({"fallback": "off", "fallback_scale": 4, "functions": []}, config.get("math", {}))
        config["images"] = merge({"raster_scale": 2, "vector": False}, config.get("images", {}))
        for kind, source in (("theme", theme_source), ("rules", rules_source)):
            _track_resource(project, _bundled(kind))
            config[kind] = str(resolve(source, kind)) if source else BUNDLED[kind][1]
        # 入口のファイルは inputs.python に書かなくても追跡する。entrypoint の表記は残し（継承元に書いたものは絶対パス）、
        # 構築は root を基準に解決する。
        entry_file = str(resolve(entrypoint.rsplit(":", 1)[0], "entrypoint"))
        inputs["python"] = [str(resolve(value, "inputs.python")) for value in inputs["python"]]
        if entry_file not in inputs["python"]:
            inputs["python"].insert(0, entry_file)
        inputs["bibliography"] = [str(resolve(value, "inputs.bibliography")) for value in inputs["bibliography"]]
        for directory in dict.fromkeys(file.parent for file in files):
            if (directory / "brief.md").is_file():
                project.brief_path = project.resolve_input(directory / "brief.md")
                break
        documents = []
        for document in inputs["documents"]:
            document = copy.deepcopy(document)
            tracked = set(project.input_files)
            document["path"] = str(resolve(document["path"], f"inputs.documents の {document.get('id')}"))
            if document.get("affects_content") is False and document["path"] not in tracked:
                # ページの内容に影響しない参考の文書は、生成記録の照合と内容の証拠から外す（存在と参照範囲は確かめる）。
                project.input_files.pop(document["path"], None)
            documents.append(document)
        inputs["documents"] = documents
        asset_dirs = []
        where["key"] = "inputs.asset_dirs"
        for value in inputs["asset_dirs"]:
            asset_dir = _local_path(value, root)
            if not asset_dir.is_dir() or not any(_within(asset_dir, base) for base in project.read_roots):
                raise ValueError(f"素材のディレクトリが存在しないか参照範囲外です: {asset_dir}")
            if (
                _within(asset_dir, output_dir)
                or _within(output_dir, asset_dir)
                or any(_within(p, asset_dir) for p in destinations)
            ):
                raise ValueError("素材と出力先の範囲が重複しています")
            asset_dirs.append(str(asset_dir))
        inputs["asset_dirs"] = asset_dirs
        where["key"] = "template"
        _template(project)
        where["key"] = ""
        if config.get("environment_override"):
            config["environment_override"] = str(resolve(config["environment_override"], "environment_override"))
        project.environment = _load_environment(project, environment, environment_paths)
    except InputError:
        raise
    except (OSError, ValueError) as error:
        message = f"の指すファイルを使えません（{error}）" if where["key"] else str(error)
        raise SettingsError([Problem(_display(path, root), where["key"], message)]) from error
    config["extends"] = [str(file) for file in files[1:]]
    return project
