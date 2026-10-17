#!/usr/bin/env python3
"""Fill a case's Juror Profiles workbook from the numbered background.

Does in one step what James did by hand each morning:

  * copies Juror Number .. household income (background columns J..X) into
    'Demographic Data' columns B..P, one row per juror (juror N -> row N+1,
    because card N's text boxes read row N+1),
  * applies the workbook's own shortening macros (ReplaceRacevalues,
    ReplaceEducationValues, ReplaceEmployValues, ReplaceMaritalValues,
    ReplaceHouseholdIncomeValues) - same words, same output,
  * puts the combined three words in column U,
  * refreshes the saved results of the card formulas and the text the card
    text boxes show, so the cards read right even before Excel recalculates,
  * on a smaller panel, deletes the unused cards and their photos and trims
    the print area,
  * sets the case name in the print header.

Only the 'Demographic Data' sheet, the card drawing and workbook.xml change;
every other part of the .xlsm (macros, pictures, card formatting) is copied
through byte-for-byte. openpyxl is never used to save it - it drops shapes.

Usage:
  python fill_profiles.py --background "Participant Background Questionnaire 1425.xlsx" \
      --profiles "1425 Juror Profiles.xlsm" --case-title "1425 Connor v. Harrow-Imprimis" \
      --out "1425 Juror Profiles.xlsm"
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from pathlib import Path

import openpyxl
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from number_jurors import (JUROR_NUM_HEADER, already_combined,  # noqa: E402
                           combine_three_words, find_three_words, norm_header)
from xlsx_package import (NS, Package, cell_text, get_cell, set_cached,  # noqa: E402
                          set_cell, shared_strings)

DATA_SHEET = "Demographic Data"
CARD_SHEET = "Juror Profiles"
CARD_TEXT_ROW0 = 34      # card N's text lives in row 34+N (A35 = juror 1)

# Demographic Data columns C..P, in the background's order (J..X there).
FIELDS = ["name", "gender", "age", "hispanic", "race", "education", "employment",
          "job", "industry", "marital", "living", "affiliation", "leaning", "income"]
FIELD_COLS = dict(zip(FIELDS, "CDEFGHIJKLMNOP"))
# A loose check that the 14 background columns are the ones we think.
FIELD_HINTS = ["your name", "gender", "age", "hispanic", "race", "education",
               "employment", "job title", "industry", "relationship", "living",
               "political affiliation", "politically", "household"]

# The workbook's shortening macros, word for word.
SHORTEN = {
    "race": {  # ReplaceRacevalues
        "American Indian or Alaska Native": "Am Indian/Alska",
        "Black or African American": "Black",
        "Native Hawaiian or Other Pacific Islander": "Hawaiian/Pacific",
        "None of these applies; I identify as Latino/Hispanic": "Hispanic",
        "Two or more race groups": "Multiracial",
    },
    "education": {  # ReplaceEducationValues
        "Bachelor’s degree": "Bachelor’s",
        "GED/High school diploma": "GED/HS",
        "Master’s degree": "Master's",
        "Associate’s degree": "Associate's",
        "Less than high school": "Less than HS",
        "Professional/Doctorate degree": "Prof/PhD",
        "Some college, no degree": "Some college",
    },
    "employment": {  # ReplaceEmployValues_ActiveSheet
        "Working full-time for employer": "Full-time/Emp.",
        "Working part-time for employer": "Part-time/Emp.",
    },
    "marital": {  # ReplaceMaritalValues
        "Single, never married": "Single",
        "In domestic partnership or civil union": "Partner/Union",
        "Cohabitating with significant other, not married or civil union": "Cohabitating",
    },
    "income": {  # ReplaceHouseholdIncomeValues
        "$15,000-$29,999": "$15k-$29,9k",
        "$30,000-$49,999": "$30k-$49,9k",
        "$50,000-$74,999": "$50k-$74,9k",
        "$75,000-$99,999": "$75k-$99,9k",
        "$100,000-$124,999": "$100k-$124,9k",
        "$125,000-$149,999": "$125k-$149,9k",
        "$150,000-$174,999": "$150k-$174,9k",
        "$175,000-$199,999": "$175k-$199,9k",
        "$200,000 or more": "$200k+",
    },
}
LONG_TEXT = 28   # job titles longer than this wrap on the card


def _apos(s: str) -> str:
    return s.replace("’", "'")


def shorten(field: str, value):
    """Exact macro replacement; apostrophe style doesn't matter for the match."""
    if not isinstance(value, str) or field not in SHORTEN:
        return value
    v = value.strip()
    for raw, short in SHORTEN[field].items():
        if _apos(v) == _apos(raw):
            return short
    return value


# ---------------------------------------------------------------- background

def read_background(path: str) -> dict[int, dict]:
    """Juror number -> {field: value, 'words': three words} from the numbered
    background (output of number_jurors.py)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    hdr = {c: norm_header(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)}
    num_col = next((c for c, h in hdr.items() if JUROR_NUM_HEADER.match(h)), None)
    name_col = next((c for c, h in hdr.items() if "what is your name" in h), None)
    if num_col is None:
        raise SystemExit("Background has no 'Juror Number' column - number the jurors first.")
    if name_col is None:
        raise SystemExit("Background has no 'What is your name?' column.")
    for i, hint in enumerate(FIELD_HINTS):
        h = hdr.get(name_col + i, "")
        if hint not in h.replace("’", "'"):
            raise SystemExit(f"Background column {name_col + i} is {h[:60]!r}; expected the "
                             f"{hint!r} question. The export's layout has changed - check it.")
    tw = find_three_words(ws)
    out = {}
    data_start = 2
    while data_start <= ws.max_row and ws.cell(data_start, 1).value in (None, ""):
        data_start += 1
    for r in range(data_start, ws.max_row + 1):
        n = ws.cell(r, num_col).value
        if n in (None, "") or not str(n).strip().isdigit():
            continue
        n = int(n)
        if n in out:
            raise SystemExit(f"Juror number {n} appears twice in the background.")
        rec = {f: ws.cell(r, name_col + i).value for i, f in enumerate(FIELDS)}
        if tw:
            w = [ws.cell(r, tw + k).value for k in range(3)]
            # number_jurors.py normally combined them into Word #1 already
            rec["words"] = str(w[0]).strip() if already_combined(*w) \
                else combine_three_words(*w)
        else:
            rec["words"] = ""
        out[n] = rec
    if not out:
        raise SystemExit("No numbered jurors in the background.")
    return out


# ---------------------------------------------------------------- card text

def _num(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return v


def _txt(v) -> str:
    v = _num(v)
    return "" if v is None else str(v)


def generation(age, table: list[tuple[float, str]]):
    """VLOOKUP(age, GenTable33, 3, TRUE). None means #N/A."""
    a = _num(age)
    if not isinstance(a, (int, float)):
        return None
    hit = None
    for lo, gen in table:
        if a >= lo:
            hit = gen
    return hit


def card_texts(rec: dict, gen, stem: str, dash: str):
    """Results of A/B/C(34+N): expanded profile, three words, condensed."""
    if gen is None:
        expanded = condensed = None          # #N/A, as Excel shows it
    else:
        expanded = (f"{stem} {_txt(rec.get('age'))} {dash} {gen}\n"
                    f"{_txt(rec.get('race'))} {dash} {_txt(rec.get('gender'))}\n"
                    f"{_txt(rec.get('job'))}\n{_txt(rec.get('education'))}\n"
                    f"{_txt(rec.get('marital'))}\n{_txt(rec.get('leaning'))}\n"
                    f"{_txt(rec.get('affiliation'))}")
        condensed = (f"{stem} {_txt(rec.get('age'))}\n{_txt(rec.get('race'))}\n"
                     f"{_txt(rec.get('education'))}\n{_txt(rec.get('job'))}\n"
                     f"{_txt(rec.get('affiliation'))}")
    words = rec.get("words") or 0           # =U2 shows 0 when U2 is empty
    return expanded, words, condensed


# ---------------------------------------------------------------- workbook

def _gen_table(sheet, sst) -> list[tuple[float, str]]:
    """Read GenTable33 (Youngest | Range | Generations) from the sheet."""
    table = []
    for r in range(1, 80):
        if cell_text(sheet, f"K{r}", sst).strip().lower() == "youngest":
            rr = r + 1
            while True:
                lo = cell_text(sheet, f"K{rr}", sst)
                gen = cell_text(sheet, f"M{rr}", sst)
                if not lo or not gen:
                    break
                table.append((float(lo), gen))
                rr += 1
            break
    if not table:
        raise SystemExit("Couldn't find the generation table (Youngest/Range/Generations).")
    return sorted(table)


def capacity(sheet) -> int:
    """How many juror rows the template has (B2, B3, ... numbered 1, 2, ...)."""
    n = 0
    while True:
        c = get_cell(sheet, f"B{n + 2}")
        v = c.findtext("m:v", default="", namespaces=NS) if c is not None else ""
        if v.strip() != str(n + 1):
            return n
        n += 1


def _anchor_name(anchor) -> str:
    el = anchor.find(".//xdr:cNvPr", NS)
    return el.get("name") if el is not None else ""


def fill(pkg: Package, jurors: dict[int, dict], case_title: str | None,
         delete_extra: bool = True) -> dict:
    data_part = pkg.sheet_part(DATA_SHEET)
    sheet = pkg.xml(data_part)
    sst = shared_strings(pkg)
    cap = capacity(sheet)
    if cap == 0:
        raise SystemExit("'Demographic Data' has no numbered juror rows (B2 = 1, B3 = 2 ...).")
    too_big = sorted(n for n in jurors if n > cap)
    if too_big:
        raise SystemExit(f"Juror number(s) {too_big} but the workbook only has {cap} cards.")
    table = _gen_table(sheet, sst)
    stem = cell_text(sheet, "Q2", sst) or "Age"
    dash = cell_text(sheet, "W2", sst) or "-"
    panel = max(jurors)
    report = {"written": [], "missing": [], "long": [], "deleted": [], "capacity": cap}
    texts = {}

    for n in range(1, cap + 1):
        r = n + 1
        rec = jurors.get(n)
        if rec:
            rec = {f: shorten(f, v) for f, v in rec.items()}
            report["written"].append(n)
            if len(_txt(rec.get("job"))) > LONG_TEXT:
                report["long"].append((n, _txt(rec.get("job"))))
        elif n <= panel:
            report["missing"].append(n)
        for f in FIELDS:
            v = rec.get(f) if rec else None
            if f == "age":
                v = _num(v)
            set_cell(sheet, f"{FIELD_COLS[f]}{r}", v)
        set_cell(sheet, f"U{r}", rec.get("words") if rec else None)
        # saved formula results, so the cards are right before a recalc
        gen = generation(rec.get("age"), table) if rec else None
        if gen is None:
            set_cached(sheet, f"S{r}", "#N/A", error=True)
        else:
            set_cached(sheet, f"S{r}", gen)
        age = _num(rec.get("age")) if rec else None
        if isinstance(age, (int, float)):
            set_cached(sheet, f"R{r}", "18-24" if 18 <= age <= 24 else
                       "25-44" if 25 <= age <= 44 else "45+")
        exp, words, cond = card_texts(rec or {}, gen, stem, dash)
        row = CARD_TEXT_ROW0 + n
        for col, val in (("A", exp), ("B", words), ("C", cond)):
            if val is None:
                set_cached(sheet, f"{col}{row}", "#N/A", error=True)
            else:
                set_cached(sheet, f"{col}{row}", val)
        texts[f"A{row}"] = "#N/A" if exp is None else exp
        texts[f"B{row}"] = str(words)
    pkg.put_xml(data_part, sheet)

    # ---- the card drawing: linked text, unused cards
    card_part = pkg.sheet_part(CARD_SHEET)
    dpart = pkg.drawing_part(card_part)
    if dpart:
        dr = pkg.xml(dpart)
        for sp in dr.iter(f"{{{NS['xdr']}}}sp"):
            link = sp.get("textlink") or ""
            m = re.match(r"^'?Demographic Data'?!\$?([A-Z]+)\$?(\d+)$", link)
            if not m or f"{m.group(1)}{m.group(2)}" not in texts:
                continue
            t = sp.find(".//a:fld/a:t", NS)
            if t is not None:
                t.text = texts[f"{m.group(1)}{m.group(2)}"]
        if delete_extra and panel < cap:
            names = {_anchor_name(a): a for a in dr}
            card_names = [f"Juror {k} Card" for k in range(1, cap + 1)]
            if not any(c in names for c in card_names):
                report["deleted"] = None   # master not tidied - can't tell the cards apart
            else:
                for k in range(panel + 1, cap + 1):
                    for nm in (f"Juror {k} Card", f"Juror {k}"):
                        if nm in names:
                            dr.remove(names[nm])
                    report["deleted"].append(k)
        pkg.put_xml(dpart, dr)

    # ---- print area and header
    wb = pkg.xml("xl/workbook.xml")
    sheets = [s.get("name") for s in wb.iterfind("m:sheets/m:sheet", NS)]
    card_idx = str(sheets.index(CARD_SHEET))
    last_card = panel if delete_extra and report["deleted"] else cap
    last_row = 5 * math.ceil(last_card / 2) + 1
    for dn in wb.iterfind("m:definedNames/m:definedName", NS):
        if dn.get("name") == "_xlnm.Print_Area" and dn.get("localSheetId") == card_idx:
            dn.text = re.sub(r"(\$[A-Z]+\$)\d+$", rf"\g<1>{last_row}", dn.text)
    pkg.put_xml("xl/workbook.xml", wb)
    pkg.full_calc_on_load()

    if case_title:
        cs = pkg.xml(card_part)
        hdr = cs.find("m:headerFooter/m:oddHeader", NS)
        if hdr is not None and "&R" in (hdr.text or ""):
            hdr.text = re.sub(r"&R(&\d+)?\s*.*$", lambda m: f"&R{m.group(1) or ''} {case_title}",
                              hdr.text)
            pkg.put_xml(card_part, cs)
            report["header"] = case_title
    return report


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Fill the Juror Profiles workbook.")
    ap.add_argument("--background", required=True, help="numbered background (.xlsx)")
    ap.add_argument("--profiles", required=True, help="the case's Juror Profiles .xlsm")
    ap.add_argument("--out", required=True, help="where to save the filled workbook")
    ap.add_argument("--case-title", help="right-hand print header, e.g. '1425 Connor v. Harrow'")
    ap.add_argument("--keep-extra-cards", action="store_true",
                    help="don't delete cards beyond the panel size")
    ap.add_argument("--overwrite", action="store_true", help="allow --out = --profiles")
    a = ap.parse_args(argv)
    if Path(a.out).resolve() == Path(a.profiles).resolve() and not a.overwrite:
        raise SystemExit("--out is the input workbook; pass --overwrite if you mean it.")

    jurors = read_background(a.background)
    pkg = Package(a.profiles)
    rep = fill(pkg, jurors, a.case_title, delete_extra=not a.keep_extra_cards)
    pkg.save(a.out)

    print(f"Filled {len(rep['written'])} jurors into {Path(a.out).name} "
          f"(workbook has {rep['capacity']} cards)")
    if rep["missing"]:
        print(f"  No background for juror(s) {rep['missing']} - their cards are blank.")
    if rep["deleted"] is None:
        print("  Extra cards NOT deleted: this workbook hasn't been tidied (no 'Juror N Card' "
              "names). Delete them by hand, or tidy the master once with tidy_profiles_master.py.")
    elif rep["deleted"]:
        print(f"  Deleted unused card(s) {rep['deleted']} and trimmed the print area.")
    if rep.get("header"):
        print(f"  Print header: {rep['header']}")
    for n, job in rep["long"]:
        print(f"  Long job title, may wrap on the card: J{n} \"{job}\"")
    print("  Open it in Excel once (it recalculates on open), add the photos, then save.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
