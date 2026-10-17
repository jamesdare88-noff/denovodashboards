#!/usr/bin/env python3
"""One-time tidy of a Juror Profiles workbook (the master, or an old case's copy).

Every card ends up built and named the same way:

    Juror N            the photo (unchanged - the photo macros and the
                       dashboard tools find photos by this name)
    Juror N Card       the card group, holding:
      Juror N Background   the card picture
      Juror Na             profile text box  (='Demographic Data'!A(34+N))
      Juror Nb             three-words box   (='Demographic Data'!B(34+N))
      Juror N Label        the "Juror N" caption
      Juror N Line         the divider line
      Juror N Card Part k  any group nested inside the card

It also:
  * groups cards that were loose pieces (10, 23 and 27 on the master). This
    matters: ResizeJPGImages and RemovePhotosFromSheet act on every loose
    picture, so a loose card background gets squashed or deleted with the
    photos,
  * deletes leftovers hidden under the cards (stray lines, an empty text box),
  * rewrites the hidden cell labels that CenterMatchingObjects uses to place
    photos - "Juror N" under each photo, "Juror Na" beside it - and clears the
    stray copies,
  * puts the Age Range formula back in every juror row of 'Demographic Data'
    and clears the broken #REF! rows under the card text,
  * sets the print header to "#### Case Name".

Formatting, text-box formulas, photos and macros are not touched. Only the
two sheets, the card drawing and calcChain change; everything else in the
.xlsm is copied through byte-for-byte.

Usage:
  python tidy_profiles_master.py --profiles "Juror Profiles Master Copy.xlsm" \
      --out "Juror Profiles Master (tidied).xlsm"
"""
from __future__ import annotations

import argparse
import copy
import re
import sys
from collections import defaultdict
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_package import (NS, Package, cell_text, get_cell, set_cell,  # noqa: E402
                          set_formula, shared_strings)

DATA_SHEET = "Demographic Data"
CARD_SHEET = "Juror Profiles"
CARD_TEXT_ROW0 = 34
X = f"{{{NS['xdr']}}}"
A_ = f"{{{NS['a']}}}"
SHAPE_TAGS = {X + t for t in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame")}
AGE_RANGE = 'IF(AND(E{r}>=18, E{r}<=24), "18-24", IF(AND(E{r}>=25, E{r}<=44), "25-44", "45+"))'
JUROR_LABEL = re.compile(r"^Juror \d+a{0,2}$")


# ---------------------------------------------------------------- drawing helpers

def shape_of(anchor):
    for ch in anchor:
        if ch.tag in SHAPE_TAGS:
            return ch
    return None


def cnvpr(el):
    for path in ("xdr:nvSpPr/xdr:cNvPr", "xdr:nvPicPr/xdr:cNvPr", "xdr:nvGrpSpPr/xdr:cNvPr",
                 "xdr:nvCxnSpPr/xdr:cNvPr", "xdr:nvGraphicFramePr/xdr:cNvPr"):
        c = el.find(path, NS)
        if c is not None:
            return c
    return None


def name_of(el) -> str:
    c = cnvpr(el)
    return c.get("name") if c is not None else ""


def rename(el, new: str) -> None:
    c = cnvpr(el)
    if c is not None:
        c.set("name", new)


def text_of(el) -> str:
    return "".join(t.text or "" for t in el.iter(A_ + "t")).strip()


def link_row(el):
    """('A'|'B', N) when the shape is a card text box linked to Demographic Data."""
    m = re.match(r"^'?Demographic Data'?!\$?([AB])\$?(\d+)$", el.get("textlink") or "")
    if m:
        return m.group(1), int(m.group(2)) - CARD_TEXT_ROW0
    return None


def anchor_pos(anchor):
    def pt(tag):
        p = anchor.find(f"xdr:{tag}", NS)
        if p is None:
            return None
        return tuple(int(p.findtext(f"xdr:{k}", namespaces=NS)) for k in ("col", "colOff", "row", "rowOff"))
    return pt("from"), pt("to")


def card_from_position(anchor) -> int | None:
    """Cards sit two per band of five rows: odd cards in B-D, even in F-G."""
    fr, to = anchor_pos(anchor)
    if fr is None or to is None:
        return None
    if fr[0] >= 8:                       # buttons and notes right of the cards
        return None
    centre = (fr[2] + to[2]) / 2
    band = int((centre - 1) // 5)
    return 2 * band + (1 if fr[0] <= 3 else 2)


def xfrm_box(el):
    """(x0, y0, x1, y1) in EMU from a top-level shape's own transform."""
    if el.tag == X + "grpSp":
        xf = el.find("xdr:grpSpPr/a:xfrm", NS)
    else:
        xf = el.find("xdr:spPr/a:xfrm", NS)
    off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
    x, y = int(off.get("x")), int(off.get("y"))
    return x, y, x + int(ext.get("cx")), y + int(ext.get("cy"))


# ---------------------------------------------------------------- naming

def name_card_parts(group, n: int, depth: int = 0, counters=None) -> None:
    counters = counters if counters is not None else defaultdict(int)
    rename(group, f"Juror {n} Card" if depth == 0 else f"Juror {n} Card Part {counters['part'] + 1}")
    if depth:
        counters["part"] += 1
    for ch in group:
        if ch.tag == X + "grpSp":
            name_card_parts(ch, n, depth + 1, counters)
        elif ch.tag == X + "sp":
            lr = link_row(ch)
            if lr:
                if lr[1] != n:
                    raise SystemExit(f"Card {n} has a text box linked to juror {lr[1]}'s row "
                                     f"({ch.get('textlink')}). Fix the link before tidying.")
                rename(ch, f"Juror {n}{lr[0].lower()}")
            elif text_of(ch):
                rename(ch, f"Juror {n} Label")
        elif ch.tag == X + "pic":
            rename(ch, f"Juror {n} Background")
        elif ch.tag == X + "cxnSp":
            counters["line"] += 1
            rename(ch, f"Juror {n} Line" if counters["line"] == 1 else f"Juror {n} Line {counters['line']}")


def make_group(drawing, anchors: list, n: int, next_id: int):
    """Wrap loose top-level shapes in one group anchor, at the z-position of
    the lowest member, keeping their order."""
    froms, tos, boxes = [], [], []
    for a in anchors:
        fr, to = anchor_pos(a)
        froms.append(fr)
        tos.append(to)
        boxes.append(xfrm_box(shape_of(a)))
    x0, y0 = min(b[0] for b in boxes), min(b[1] for b in boxes)
    x1, y1 = max(b[2] for b in boxes), max(b[3] for b in boxes)
    fc = min((f[0], f[1]) for f in froms)
    fr_ = min((f[2], f[3]) for f in froms)
    tc = max((t[0], t[1]) for t in tos)
    tr = max((t[2], t[3]) for t in tos)

    E = lambda tag, parent=None, **kw: (etree.SubElement(parent, tag, **kw)  # noqa: E731
                                        if parent is not None else etree.Element(tag, **kw))
    anchor = E(X + "twoCellAnchor")
    for tag, (c, co), (r, ro) in (("from", fc, fr_), ("to", tc, tr)):
        p = E(X + tag, anchor)
        E(X + "col", p).text = str(c)
        E(X + "colOff", p).text = str(co)
        E(X + "row", p).text = str(r)
        E(X + "rowOff", p).text = str(ro)
    g = E(X + "grpSp", anchor)
    nv = E(X + "nvGrpSpPr", g)
    E(X + "cNvPr", nv, id=str(next_id), name=f"Juror {n} Card")
    E(X + "cNvGrpSpPr", nv)
    xf = E(A_ + "xfrm", E(X + "grpSpPr", g))
    for tag, (a, b) in (("off", (x0, y0)), ("ext", (x1 - x0, y1 - y0)),
                        ("chOff", (x0, y0)), ("chExt", (x1 - x0, y1 - y0))):
        if tag.endswith("ff"):
            E(A_ + tag, xf, x=str(a), y=str(b))
        else:
            E(A_ + tag, xf, cx=str(a), cy=str(b))
    for a in anchors:
        g.append(copy.deepcopy(shape_of(a)))
    E(X + "clientData", anchor)
    anchors[0].addprevious(anchor)
    for a in anchors:
        drawing.remove(a)
    return g


def tidy_drawing(drawing) -> dict:
    rep = {"grouped": [], "deleted": [], "renamed": 0, "problems": []}
    ids = [int(c.get("id")) for c in drawing.iter(X + "cNvPr") if c.get("id", "").isdigit()]
    next_id = max(ids, default=1) + 1
    before = {name_of(e) for e in drawing.iter() if e.tag in SHAPE_TAGS}

    per_card = defaultdict(lambda: {"photo": None, "groups": [], "loose": []})
    for a in list(drawing):
        el = shape_of(a)
        if el is None:
            continue
        nm = name_of(el)
        if nm.startswith("Button") or el.tag == X + "graphicFrame":
            continue
        m = re.match(r"^Juror (\d+)$", nm)
        if el.tag == X + "pic" and m:
            per_card[int(m.group(1))]["photo"] = a
            continue
        lr = link_row(el) if el.tag == X + "sp" else None
        n = lr[1] if lr else card_from_position(a)
        if n is None:
            rep["problems"].append(f"Left alone (not on a card): {nm}")
            continue
        per_card[n]["groups" if el.tag == X + "grpSp" else "loose"].append(a)

    for n in sorted(per_card):
        c = per_card[n]
        if len(c["groups"]) > 1:
            rep["problems"].append(f"Card {n} has {len(c['groups'])} groups - left as they are")
            continue
        loose = c["loose"]
        if c["groups"]:
            # a grouped card: anything loose on it is a hidden leftover
            for a in loose:
                el = shape_of(a)
                if el.tag == X + "cxnSp" or (el.tag == X + "sp" and not text_of(el)
                                             and not link_row(el)):
                    rep["deleted"].append(f"Juror {n}: {name_of(el)}")
                    drawing.remove(a)
                else:
                    rep["problems"].append(f"Card {n}: loose {name_of(el)} next to the card group "
                                           f"- left as it is")
            group = shape_of(c["groups"][0])
        else:
            keep = []
            for a in loose:
                el = shape_of(a)
                if el.tag == X + "sp" and not text_of(el) and not link_row(el):
                    rep["deleted"].append(f"Juror {n}: {name_of(el)}")
                    drawing.remove(a)
                else:
                    keep.append(a)
            if not keep:
                continue
            group = make_group(drawing, keep, n, next_id)
            next_id += 1
            rep["grouped"].append(n)
        name_card_parts(group, n)

    # check every card came out complete
    names = defaultdict(list)
    for e in drawing.iter():
        if e.tag in SHAPE_TAGS:
            names[name_of(e)].append(e)
    for n in sorted(k for k in per_card if per_card[k]["photo"] is not None or per_card[k]["groups"]
                    or per_card[k]["loose"]):
        for need in (f"Juror {n}", f"Juror {n} Card", f"Juror {n}a", f"Juror {n}b",
                     f"Juror {n} Background", f"Juror {n} Label", f"Juror {n} Line"):
            if len(names.get(need, [])) != 1:
                rep["problems"].append(f"Card {n}: expected one '{need}', found "
                                       f"{len(names.get(need, []))}")
    after = {name_of(e) for e in drawing.iter() if e.tag in SHAPE_TAGS}
    rep["renamed"] = len(after - before)
    return rep


# ---------------------------------------------------------------- sheets

def tidy_labels(sheet, sst, cards: list[int]) -> list[str]:
    """Hidden labels under the cards: 'Juror N' under each photo, 'Juror Na'
    beside it; every other 'Juror ...' label cleared."""
    want = {}
    for n in cards:
        band, left = (n - 1) // 2, n % 2 == 1
        r = 4 + 5 * band
        want[f"{'C' if left else 'F'}{r}"] = f"Juror {n}"
        want[f"{'D' if left else 'G'}{r}"] = f"Juror {n}a"
    changes = []
    style = {}
    for row in sheet.find("m:sheetData", NS):
        for c in row:
            ref = c.get("r")
            col = re.match(r"[A-Z]+", ref).group(0)
            if c.get("s"):
                style.setdefault(col, c.get("s"))
            txt = cell_text(sheet, ref, sst)
            if JUROR_LABEL.match(txt) and want.get(ref) != txt:
                if ref not in want:
                    changes.append(f"cleared {ref} '{txt}'")
                    set_cell(sheet, ref, None)
    for ref, txt in want.items():
        if cell_text(sheet, ref, sst) != txt:
            changes.append(f"{ref} -> '{txt}'")
            set_cell(sheet, ref, txt)
            c = get_cell(sheet, ref)
            col = re.match(r"[A-Z]+", ref).group(0)
            if not c.get("s") and col in style:
                c.set("s", style[col])
    return changes


def tidy_data_sheet(sheet, sst, cap: int) -> list[str]:
    changes = []
    for r in range(2, cap + 2):
        c = get_cell(sheet, f"R{r}")
        if c is None or c.find("m:f", NS) is None:
            style = c.get("s") if c is not None else None
            set_formula(sheet, f"R{r}", AGE_RANGE.format(r=r))
            if style:
                get_cell(sheet, f"R{r}").set("s", style)
            changes.append(f"R{r} Age Range formula")
    # broken rows below the card text (formulas pointing at #REF! / past the table)
    last = CARD_TEXT_ROW0 + cap + 1
    for row in list(sheet.find("m:sheetData", NS)):
        r = int(row.get("r"))
        if r <= last:
            continue
        for c in row:
            f = c.find("m:f", NS)
            if f is not None and ("#REF!" in (f.text or "") or c.findtext("m:v", "", NS) == "#REF!"
                                  or re.fullmatch(r"U\d+", (f.text or "").strip() or "-")):
                set_formula(sheet, c.get("r"), None)
                changes.append(f"cleared broken {c.get('r')}")
    return changes


def tidy(pkg: Package) -> dict:
    sst = shared_strings(pkg)
    card_part = pkg.sheet_part(CARD_SHEET)
    dpart = pkg.drawing_part(card_part)
    drawing = pkg.xml(dpart)
    rep = tidy_drawing(drawing)
    pkg.put_xml(dpart, drawing)
    cards = sorted({int(m.group(1)) for e in drawing.iter() if e.tag in SHAPE_TAGS
                    for m in [re.match(r"^Juror (\d+) Card$", name_of(e))] if m})

    sheet = pkg.xml(card_part)
    rep["labels"] = tidy_labels(sheet, sst, cards)
    hdr = sheet.find("m:headerFooter/m:oddHeader", NS)
    if hdr is not None and "&R" in (hdr.text or ""):
        hdr.text = re.sub(r"&R(&\d+)?\s*.*$", lambda m: f"&R{m.group(1) or ''} #### Case Name",
                          hdr.text)
    pkg.put_xml(card_part, sheet)

    data_part = pkg.sheet_part(DATA_SHEET)
    ds = pkg.xml(data_part)
    rep["data"] = tidy_data_sheet(ds, sst, len(cards))
    pkg.put_xml(data_part, ds)
    pkg.drop_calc_chain()
    pkg.full_calc_on_load()       # the new Age Range formulas have no saved result yet
    rep["cards"] = cards
    return rep


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Tidy a Juror Profiles workbook once.")
    ap.add_argument("--profiles", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    if Path(a.out).resolve() == Path(a.profiles).resolve():
        raise SystemExit("Write the tidied copy to a new file; keep the original as a backup.")
    pkg = Package(a.profiles)
    rep = tidy(pkg)
    pkg.save(a.out)
    print(f"Tidied {len(rep['cards'])} cards -> {a.out}")
    print(f"  Grouped loose cards: {rep['grouped'] or 'none'}")
    print(f"  Deleted hidden leftovers: {rep['deleted'] or 'none'}")
    print(f"  Shapes renamed: {rep['renamed']}")
    print(f"  Cell labels fixed: {len(rep['labels'])}")
    print(f"  Demographic Data fixes: {len(rep['data'])}")
    for p in rep["problems"]:
        print(f"  CHECK: {p}")
    return 1 if rep["problems"] else 0


if __name__ == "__main__":
    sys.exit(main())
