"""Edit one or two parts of an .xlsx/.xlsm without touching the rest.

openpyxl rewrites the whole workbook and silently drops text boxes, grouped
shapes and form buttons - fatal for the Juror Profiles workbook, whose cards
are nothing but those. This class unzips the package, lets the caller swap a
few XML parts, and writes every other part back byte-for-byte, in the same
order and with the same compression. Macros (vbaProject.bin), pictures and
drawings that aren't edited come through untouched.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}


class Package:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.infos = list(z.infolist())
            self.data = {i.filename: z.read(i.filename) for i in self.infos}
        self.changed: set[str] = set()

    # ---------------------------------------------------------------- parts
    def xml(self, name: str):
        return etree.fromstring(self.data[name])

    def put_xml(self, name: str, root) -> None:
        self.data[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                         standalone=True)
        self.changed.add(name)

    def delete(self, name: str) -> None:
        self.data.pop(name, None)
        self.infos = [i for i in self.infos if i.filename != name]
        self.changed.add(name)

    def save(self, out: str | Path) -> None:
        out = Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "w") as z:
            for info in self.infos:
                z.writestr(info, self.data[info.filename], compress_type=info.compress_type)

    # ---------------------------------------------------------------- lookups
    def sheet_part(self, sheet_name: str) -> str:
        """xl/worksheets/sheetN.xml for a sheet's display name."""
        wb = self.xml("xl/workbook.xml")
        rels = self.xml("xl/_rels/workbook.xml.rels")
        for s in wb.iterfind("m:sheets/m:sheet", NS):
            if s.get("name") == sheet_name:
                rid = s.get(f"{{{NS['r']}}}id")
                for rel in rels:
                    if rel.get("Id") == rid:
                        return "xl/" + rel.get("Target").lstrip("/").removeprefix("xl/")
        raise KeyError(f"No sheet named {sheet_name!r}")

    def drawing_part(self, sheet_part: str) -> str | None:
        rel_name = re.sub(r"worksheets/(sheet\d+\.xml)$", r"worksheets/_rels/\1.rels", sheet_part)
        if rel_name not in self.data:
            return None
        for rel in self.xml(rel_name):
            if rel.get("Type", "").endswith("/drawing"):
                return str(Path(sheet_part).parent.joinpath(rel.get("Target")).as_posix()) \
                    .replace("worksheets/../", "")
        return None

    def drop_calc_chain(self) -> None:
        """Excel rebuilds calcChain.xml silently; a stale one triggers a
        'repair' prompt after formulas are added or removed."""
        if "xl/calcChain.xml" not in self.data:
            return
        self.delete("xl/calcChain.xml")
        rels = self.xml("xl/_rels/workbook.xml.rels")
        for rel in list(rels):
            if rel.get("Type", "").endswith("/calcChain"):
                rels.remove(rel)
        self.put_xml("xl/_rels/workbook.xml.rels", rels)
        ct = self.xml("[Content_Types].xml")
        for o in list(ct):
            if o.get("PartName") == "/xl/calcChain.xml":
                ct.remove(o)
        self.put_xml("[Content_Types].xml", ct)

    def full_calc_on_load(self) -> None:
        wb = self.xml("xl/workbook.xml")
        cp = wb.find("m:calcPr", NS)
        if cp is None:
            cp = etree.SubElement(wb, f"{{{NS['m']}}}calcPr")
        cp.set("fullCalcOnLoad", "1")
        self.put_xml("xl/workbook.xml", wb)


# ---------------------------------------------------------------- cells

def col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def _row(sheet, r: int):
    data = sheet.find("m:sheetData", NS)
    for row in data:
        n = int(row.get("r"))
        if n == r:
            return row
        if n > r:
            new = etree.Element(f"{{{NS['m']}}}row", r=str(r))
            row.addprevious(new)
            return new
    return etree.SubElement(data, f"{{{NS['m']}}}row", r=str(r))


def get_cell(sheet, ref: str):
    """The <c> element for ref, or None. Never adds rows or cells."""
    col, r = split_ref(ref)
    for row in sheet.find("m:sheetData", NS):
        if int(row.get("r")) == r:
            for c in row:
                if c.get("r") == ref:
                    return c
            return None
    return None


def _cell(sheet, ref: str):
    col, r = split_ref(ref)
    row = _row(sheet, r)
    for c in row:
        if c.get("r") == ref:
            return c
        if col_index(split_ref(c.get("r"))[0]) > col_index(col):
            new = etree.Element(f"{{{NS['m']}}}c", r=ref)
            c.addprevious(new)
            return new
    return etree.SubElement(row, f"{{{NS['m']}}}c", r=ref)


def set_cell(sheet, ref: str, value) -> None:
    """Write a plain value, keeping the cell's style. None/"" empties it."""
    c = _cell(sheet, ref)
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop("t", None)
    if value is None or value == "":
        return
    if isinstance(value, bool):
        value = str(value)
    if isinstance(value, (int, float)):
        etree.SubElement(c, f"{{{NS['m']}}}v").text = repr(value) if isinstance(value, float) \
            else str(value)
        return
    c.set("t", "inlineStr")
    is_ = etree.SubElement(c, f"{{{NS['m']}}}is")
    t = etree.SubElement(is_, f"{{{NS['m']}}}t")
    t.text = str(value)
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def set_cached(sheet, ref: str, value, error: bool = False) -> None:
    """Refresh a formula cell's saved result without touching the formula."""
    c = get_cell(sheet, ref)
    if c is None or c.find("m:f", NS) is None:
        return
    v = c.find("m:v", NS)
    if v is None:
        v = etree.SubElement(c, f"{{{NS['m']}}}v")
    if error:
        c.set("t", "e")
        v.text = str(value)
    elif isinstance(value, (int, float)):
        c.attrib.pop("t", None)
        v.text = str(value)
    else:
        c.set("t", "str")
        v.text = str(value)


def set_formula(sheet, ref: str, formula: str | None) -> None:
    """Put a formula in a cell (or None to clear the cell), keeping style."""
    c = _cell(sheet, ref)
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop("t", None)
    if formula:
        etree.SubElement(c, f"{{{NS['m']}}}f").text = formula


def shared_strings(pkg: Package) -> list[str]:
    if "xl/sharedStrings.xml" not in pkg.data:
        return []
    root = pkg.xml("xl/sharedStrings.xml")
    return ["".join(si.itertext()) for si in root.iterfind("m:si", NS)]


def cell_text(sheet, ref: str, sst: list[str]) -> str:
    c = get_cell(sheet, ref)
    if c is None:
        return ""
    if c.get("t") == "s":
        return sst[int(c.findtext("m:v", namespaces=NS))]
    if c.get("t") == "inlineStr":
        return "".join(c.find("m:is", NS).itertext())
    return c.findtext("m:v", default="", namespaces=NS)
