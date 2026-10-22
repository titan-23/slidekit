"""OOXML の名前空間と、PPTX のパッケージの読み書きの補助。"""

from __future__ import annotations

from pathlib import Path
import posixpath
import zipfile

from lxml import etree

from .files import atomic_output

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "dc": "http://purl.org/dc/elements/1.1/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
}
EMU_PER_PT = 12700


def tag(prefix: str, local: str) -> str:
    """接頭辞と局所名から、名前空間付きの要素名を作る。"""
    return f"{{{NS[prefix]}}}{local}"


def read_package(path: str | Path) -> dict[str, bytes]:
    """PPTX の部品を名前から中身への辞書で読む。破損した部品があれば ValueError にする。"""
    with zipfile.ZipFile(path) as package:
        corrupt = package.testzip()
        if corrupt:
            raise ValueError(f"壊れた PPTX の要素: {corrupt}")
        return {name: package.read(name) for name in package.namelist()}


def write_package(parts: dict[str, bytes], output: str | Path) -> None:
    """部品の辞書を PPTX（ZIP）として原子的に書く。"""
    with atomic_output(output) as stream:
        with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for name, contents in parts.items():
                package.writestr(name, contents)


def xml_bytes(root) -> bytes:
    """XML 宣言付きの UTF-8 のバイト列にする。"""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def slide_order(parts: dict[str, bytes]) -> tuple[list[str], tuple[float, float]]:
    """スライドの部品名を表示順に並べ、ページ寸法（pt）と合わせて返す。"""
    root = etree.fromstring(parts["ppt/presentation.xml"])
    relationships = etree.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    targets = {node.get("Id"): node.get("Target") for node in relationships}
    ordered = [
        posixpath.normpath(posixpath.join("ppt", targets[node.get(tag("r", "id"))])).lstrip("/")
        for node in root.findall("p:sldIdLst/p:sldId", NS)
    ]
    size = root.find("p:sldSz", NS)
    return ordered, (int(size.get("cx")) / EMU_PER_PT, int(size.get("cy")) / EMU_PER_PT)
