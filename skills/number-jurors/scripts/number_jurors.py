#!/usr/bin/env python3
"""Number the jurors in a SurveyMonkey Participant Background Questionnaire.

At morning check-in each seated juror gets a juror number, recorded in the
recruiter's sign-in workbook (on Nelson sheets: the "J#" column of "Amys Tab").
This script copies those numbers into the SurveyMonkey background export:

  * the `rq_flag` column becomes `Juror Number` (a column is inserted if the
    export has neither),
  * each seated juror's row gets their number, matched by name,
  * rows are sorted by juror number, whole rows moving together,
  * respondents who were not seated (no-shows, alternates, partial or
    duplicate submissions) are removed.

It never guesses. A seated juror whose name can't be matched with confidence
stops the run (exit code 2) and the report lists the closest candidates, so a
person can confirm the pairing and re-run with --map.

Usage:
  python number_jurors.py --background "Participant Background Questionnaire 1425.xlsx" \
      --signin "Shared Philadelphia, PA Thursday October 1st.xlsx" \
      --out "numbered/Participant Background Questionnaire 1425.xlsx"

  # numbers typed by hand instead of a sign-in workbook ("1 Jane Doe" per line):
  python number_jurors.py --background BG.xlsx --roster roster.txt --out OUT.xlsx

  # confirm a pairing the script could not make on its own:
  ... --map "20=Shannon" --map "13=Stephan Rodriguez"

  # a seated juror who never filled out the background:
  ... --allow-missing 7
"""
from __future__ import annotations

import argparse
import difflib
import re
import sys
import unicodedata
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- headers

# "J#", "Juror #", "Juror Number", "Juror Numbers", "Juror No.", "Jur #".
# A bare "#" is deliberately NOT accepted: on recruiter sheets that is the
# recruit/seat number, not the juror number.
JUROR_NUM_HEADER = re.compile(
    r"^\s*(j\s*#|jur(or)?\s*(#|no\.?|nos\.?|num(ber)?s?))\s*$", re.I)
NAME_HEADERS = ("rec names", "rec name", "juror name", "juror names", "full name",
                "name", "names", "participant", "participant name")
BG_NAME_HEADER = re.compile(r"what is your name", re.I)
NUM_VALUE = re.compile(r"^\s*(?:j|juror)?\s*#?\s*(\d{1,3})(?![\d.])\s*[-:,]?\s*(.*?)\s*$", re.I)
JUNK_NAME = re.compile(r"^column\d+$", re.I)  # Excel-table artefact rows

# Words recruiters append to a name that are not part of it.
STATUS_WORDS = {"late", "eb", "absent", "noshow", "alt", "alternate", "standby"}
SUFFIXES = {"jr", "sr", "ii", "iii", "iv"}


def norm_header(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip().lower() if v is not None else ""


# ---------------------------------------------------------------- names

def _ascii(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))


def _tokens(s: str) -> list[str]:
    s = _ascii(s).lower().replace("’", "'")
    s = re.sub(r"[^a-z' -]", " ", s)
    s = s.replace("'", "")
    toks = [t for t in re.split(r"[\s]+", s) if t and t != "-"]
    toks = [t for t in toks if t not in SUFFIXES]
    while toks and toks[-1] in STATUS_WORDS:
        toks.pop()
    return toks


def name_variants(raw) -> list[list[str]]:
    """Token lists a name could be known by. The first entry is the name as
    written; later entries swap in any nickname given in quotes/brackets."""
    if raw is None:
        return []
    s = str(raw).replace("''", '"')
    s = re.split(r"\s+[-–—]\s+", s)[0]            # "Morgan Swanson - released at lunch"
    s = re.sub(r"\*+\s*[A-Za-z]*", " ", s)         # "***EB"
    nicks = re.findall(r"[\"“”]([^\"“”]+)[\"“”]|\(([^)]*)\)", s)
    nicks = [a or b for a, b in nicks]
    base = re.sub(r"[\"“”][^\"“”]*[\"“”]|\([^)]*\)", " ", s)
    out = []
    bt = _tokens(base)
    if bt:
        out.append(bt)
    for n in nicks:
        nt = _tokens(n)
        if len(nt) == 1 and len(bt) >= 2:
            out.append(nt + bt[1:])
        elif len(nt) >= 2:
            out.append(nt)
    return out


def _drop_initials(t: list[str]) -> list[str]:
    return [x for i, x in enumerate(t) if not (0 < i < len(t) - 1 and len(x) == 1)] or t


# ---------------------------------------------------------------- inputs

@dataclass
class Seat:
    num: int
    name: str
    variants: list = field(default_factory=list)
    note: str = ""        # text written next to the number, e.g. "21 MIA-CUT"
    attendance: str = ""


def parse_num(v):
    """Juror number from a J# cell -> (number, note). A cell with no leading
    number ("GHOSTED", blank) means not seated -> (None, "")."""
    if v is None or isinstance(v, bool):
        return None, ""
    if isinstance(v, (int, float)):
        return (int(v), "") if float(v).is_integer() and v > 0 else (None, "")
    m = NUM_VALUE.match(str(v))
    if not m or int(m.group(1)) == 0:
        return None, ""
    return int(m.group(1)), m.group(2).strip()


def _find_name_col(hdrs: dict[int, str]):
    for want in NAME_HEADERS:
        for c, h in hdrs.items():
            if h == want:
                return ("single", c)
    first = next((c for c, h in hdrs.items() if h in ("first name", "first")), None)
    last = next((c for c, h in hdrs.items() if h in ("last name", "last")), None)
    if first and last:
        return ("pair", first, last)
    return None


def read_signin(path: str, sheet: str | None = None) -> tuple[list[Seat], str]:
    """Find the juror-number column in the sign-in workbook and return the
    seated jurors. Every sheet is scanned unless `sheet` is given; the table
    with the most numbered rows wins."""
    wb = openpyxl.load_workbook(path, data_only=True)
    wbf = None  # formula view, loaded only if a name cell has no cached value

    def cell_value(ws, r, c):
        nonlocal wbf
        v = ws.cell(r, c).value
        if v is not None:
            return v
        if wbf is None:
            wbf = openpyxl.load_workbook(path)
        f = wbf[ws.title].cell(r, c).value
        m = re.match(r"^=\s*(?:'([^']+)'|([A-Za-z0-9_ ]+))!\$?([A-Z]+)\$?(\d+)\s*$", str(f or ""))
        if m:  # simple ='Sheet'!G5 reference that Excel never calculated
            src = wb[m.group(1) or m.group(2)]
            return src[f"{m.group(3)}{m.group(4)}"].value
        return None

    sheets = [wb[sheet]] if sheet else wb.worksheets
    found = []
    for ws in sheets:
        for r in range(1, min(ws.max_row, 20) + 1):
            hdrs = {c: norm_header(ws.cell(r, c).value)
                    for c in range(1, min(ws.max_column, 80) + 1)}
            num_cols = [c for c, h in hdrs.items() if h and JUROR_NUM_HEADER.match(h)]
            namecol = _find_name_col(hdrs)
            if not num_cols or not namecol:
                continue
            nc = num_cols[0]
            att = next((c for c, h in hdrs.items() if h == "attendance"), None)
            seats = []
            for rr in range(r + 1, ws.max_row + 1):
                n, note = parse_num(ws.cell(rr, nc).value)
                if n is None:
                    continue
                if namecol[0] == "single":
                    nm = cell_value(ws, rr, namecol[1])
                else:
                    nm = " ".join(str(cell_value(ws, rr, c) or "").strip()
                                  for c in namecol[1:]).strip()
                if nm is None or not str(nm).strip():
                    continue
                a = str(cell_value(ws, rr, att) or "").strip() if att else ""
                seats.append(Seat(n, str(nm).strip(), name_variants(nm), note, a))
            if seats:
                found.append((len(seats), f"{ws.title}!{get_column_letter(nc)}", seats))
            break  # only the first header row per sheet
    if not found:
        raise SystemExit(
            f"No juror-number column (J#, Juror #, Juror Number) with a name column "
            f"found in {path}. Pass --sheet, or give the numbers with --roster.")
    found.sort(key=lambda x: -x[0])
    return found[0][2], found[0][1]


def read_roster(path: str) -> list[Seat]:
    seats = []
    for line in Path(path).read_text(encoding="utf-8-sig").splitlines():
        m = re.match(r"^\s*(?:j|juror)?\s*#?\s*(\d{1,3})\s*[.):,\t-]?\s+(.+?)\s*$", line, re.I) \
            or re.match(r"^\s*(\d{1,3})\s*,\s*(.+?)\s*$", line)
        if m:
            nm = m.group(2).strip().strip(",").strip('"')
            seats.append(Seat(int(m.group(1)), nm, name_variants(nm)))
    if not seats:
        raise SystemExit(f"No '<number> <name>' lines found in {path}")
    return seats


# ---------------------------------------------------------------- background

@dataclass
class Respondent:
    row: int              # original sheet row
    name: str
    variants: list
    filled: int           # non-empty answer cells, for picking among duplicates
    rid: str


@dataclass
class Background:
    wb: object
    ws: object
    data_start: int
    name_col: int
    num_col: int | None   # None -> must insert
    num_header: str       # header text found in the num column (or "")
    respondents: list
    skipped: list         # (row, reason)


def load_background(path: str) -> Background:
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    hdr = {c: norm_header(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)}
    name_col = next((c for c, h in hdr.items() if BG_NAME_HEADER.search(h)), None)
    if name_col is None:
        raise SystemExit("Background export has no 'What is your name?' column in row 1.")
    # An existing "Juror Number(s)" column wins (1394 keeps it in column I with
    # rq_flag still in J); otherwise rq_flag is the column de novo repurposes.
    num_col = next((c for c, h in hdr.items() if JUROR_NUM_HEADER.match(h)), None) or \
        next((c for c, h in hdr.items() if h == "rq_flag"), None)
    # Row 2 of a SurveyMonkey export is the Response / Open-Ended Response row.
    data_start = 2
    while data_start <= ws.max_row and ws.cell(data_start, 1).value in (None, ""):
        data_start += 1
    resp, skipped = [], []
    for r in range(data_start, ws.max_row + 1):
        nm = ws.cell(r, name_col).value
        if nm is None or not str(nm).strip():
            if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                skipped.append((r, "no name (partial submission)"))
            continue
        if JUNK_NAME.match(str(nm).strip()):
            skipped.append((r, f"artefact row ('{nm}')"))
            continue
        filled = sum(ws.cell(r, c).value not in (None, "")
                     for c in range(name_col, ws.max_column + 1))
        resp.append(Respondent(r, str(nm).strip(), name_variants(nm), filled,
                               str(ws.cell(r, 1).value or "")))
    return Background(wb, ws, data_start, name_col, num_col,
                      hdr.get(num_col, "") if num_col else "", resp, skipped)


# ---------------------------------------------------------------- matching

TIER_LABEL = {1: "exact", 2: "first + last", 3: "similar spelling",
              4: "first name only", 5: "close spelling", 9: "confirmed (--map)"}


@dataclass
class MatchResult:
    pairs: dict           # juror num -> (Respondent, tier)
    unmatched: list       # Seat
    ambiguous: dict       # juror num -> [Respondent]
    duplicates: list      # Respondent rows dropped as duplicate submissions


def _similar(a: str, b: str) -> bool:
    return a.startswith(b) or b.startswith(a) or \
        difflib.SequenceMatcher(None, a, b).ratio() >= 0.8


def _last_same(a: str, b: str) -> bool:
    """Surnames agree, allowing for one side's hyphenated double surname
    ("Thompson-Matthews" / "Thompson", "Shanta-Chano" / "Chano")."""
    pa, pb = set(a.split("-")), set(b.split("-"))
    return a == b or bool(pa & pb)


def _candidates(seat: Seat, pool: list[Respondent], tier: int, all_seats: list[Seat]):
    out = []
    for r in pool:
        hit = False
        for sv in seat.variants:
            s2 = _drop_initials(sv)
            for rv in r.variants:
                r2 = _drop_initials(rv)
                if tier == 1:
                    hit = sv == rv or s2 == r2
                elif tier == 2:
                    hit = len(s2) >= 2 and len(r2) >= 2 and s2[0] == r2[0] and \
                        _last_same(s2[-1], r2[-1])
                    # sign-in name followed by notes: "Becky Morales was online but..."
                    hit = hit or (len(r2) >= 2 and s2[:len(r2)] == r2)
                elif tier == 3:
                    # one side misspelt: first OR last name the same, the other close
                    hit = len(s2) >= 2 and len(r2) >= 2 and (
                        (_last_same(s2[-1], r2[-1]) and _similar(s2[0], r2[0])) or
                        (s2[0] == r2[0] and _similar(s2[-1], r2[-1])))
                elif tier == 4:
                    # respondent typed only a first name; it must be unique among
                    # seated jurors too, or we can't tell who it was.
                    hit = len(r2) == 1 and s2 and s2[0] == r2[0] and \
                        sum(1 for o in all_seats if o.variants and
                            _drop_initials(o.variants[0])[0] == r2[0]) == 1
                if hit:
                    break
            if hit:
                break
        if hit:
            out.append(r)
    if tier == 5:
        scored = []
        for r in pool:
            best = max((difflib.SequenceMatcher(None, " ".join(sv), " ".join(rv)).ratio()
                        for sv in seat.variants for rv in r.variants), default=0)
            scored.append((best, r))
        scored.sort(key=lambda x: -x[0])
        if scored and scored[0][0] >= 0.88 and \
                (len(scored) == 1 or scored[0][0] - scored[1][0] >= 0.06):
            out = [scored[0][1]]
    return out


def dedupe(respondents: list[Respondent]):
    """Same person submitted twice: keep the most complete submission."""
    groups = {}
    for r in respondents:
        key = " ".join(r.variants[0]) if r.variants else r.name.lower()
        groups.setdefault(key, []).append(r)
    keep, dropped = [], []
    for g in groups.values():
        g = sorted(g, key=lambda r: (-r.filled, r.row))
        keep.append(g[0])
        dropped.extend(g[1:])
    keep.sort(key=lambda r: r.row)
    return keep, dropped


def match(seats: list[Seat], respondents: list[Respondent],
          manual: dict[int, str] | None = None) -> MatchResult:
    pool, dupes = dedupe(respondents)
    pairs, ambiguous = {}, {}
    seats_left = {s.num: s for s in seats}

    for num, target in (manual or {}).items():
        if num not in seats_left:
            raise SystemExit(f"--map {num}=…: juror {num} is not on the seated list.")
        t = target.strip().lower()
        hits = [r for r in pool if r.name.strip().lower() == t or r.rid == target.strip()]
        if len(hits) != 1:
            raise SystemExit(f"--map {num}={target}: matched {len(hits)} background rows; "
                             f"use the exact name or Respondent ID.")
        pairs[num] = (hits[0], 9)
        pool.remove(hits[0])
        del seats_left[num]

    for tier in (1, 2, 3, 4, 5):
        claims = {}
        for num, seat in sorted(seats_left.items()):
            if num in ambiguous:
                continue
            c = _candidates(seat, pool, tier, seats)
            if len(c) == 1:
                claims.setdefault(id(c[0]), []).append((num, c[0]))
            elif len(c) > 1:
                ambiguous[num] = c
        for lst in claims.values():
            if len(lst) == 1:                    # one seat, one row: take it
                num, r = lst[0]
                pairs[num] = (r, tier)
                pool.remove(r)
                del seats_left[num]
            else:                                # two seats want one row
                for num, r in lst:
                    ambiguous[num] = [r]
    unmatched = [s for n, s in sorted(seats_left.items()) if n not in ambiguous]
    return MatchResult(pairs, unmatched, ambiguous, dupes)


def suggestions(seat: Seat, pool: list[Respondent], k: int = 3):
    scored = []
    for r in pool:
        best = max((difflib.SequenceMatcher(None, " ".join(sv), " ".join(rv)).ratio()
                    for sv in seat.variants for rv in r.variants), default=0)
        scored.append((best, r.name))
    return [n for _, n in sorted(scored, reverse=True)[:k]]


# ---------------------------------------------------------------- writing

def _insert_col(ws, idx: int):
    """insert_cols that also shifts merged ranges (openpyxl leaves them)."""
    merged = [str(m) for m in ws.merged_cells.ranges]
    for m in merged:
        ws.unmerge_cells(m)
    ws.insert_cols(idx)
    for m in merged:
        a, b = m.split(":") if ":" in m else (m, m)
        ca = openpyxl.utils.cell.coordinate_from_string(a)
        cb = openpyxl.utils.cell.coordinate_from_string(b)
        c1 = openpyxl.utils.column_index_from_string(ca[0])
        c2 = openpyxl.utils.column_index_from_string(cb[0])
        if c1 >= idx:
            c1 += 1
        if c2 >= idx:
            c2 += 1
        ws.merge_cells(start_row=ca[1], start_column=c1, end_row=cb[1], end_column=c2)


def write_numbered(bg: Background, pairs: dict, out: str):
    ws = bg.ws
    if bg.num_col is None:
        # put it where de novo keeps it: right after "Custom Data 1", else
        # immediately before the name question.
        hdr = {c: norm_header(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)}
        cd = next((c for c, h in hdr.items() if h.startswith("custom data")), None)
        bg.num_col = (cd + 1) if cd else bg.name_col
        _insert_col(ws, bg.num_col)
        if bg.name_col >= bg.num_col:
            bg.name_col += 1
    ws.cell(1, bg.num_col).value = "Juror Number"
    if bg.data_start > 2:
        ws.cell(2, bg.num_col).value = None

    ncols = ws.max_column
    snap = {}
    for num, (r, _) in pairs.items():
        row = r.row
        snap[num] = [(ws.cell(row, c).value, ws.cell(row, c).number_format,
                      copy(ws.cell(row, c).font), copy(ws.cell(row, c).alignment))
                     for c in range(1, ncols + 1)]
    ws.delete_rows(bg.data_start, ws.max_row - bg.data_start + 1)
    for i, num in enumerate(sorted(snap)):
        rr = bg.data_start + i
        for c, (v, fmt, font, align) in enumerate(snap[num], start=1):
            cell = ws.cell(rr, c)
            cell.value = v
            cell.number_format = fmt
            cell.font = font
            cell.alignment = align
        ws.cell(rr, bg.num_col).value = num
        ws.cell(rr, bg.num_col).number_format = "General"
    # A leftover table/filter range would point at deleted rows.
    ws.auto_filter.ref = None
    for t in list(getattr(ws, "tables", {}).keys()):
        del ws.tables[t]
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    bg.wb.save(out)


# ---------------------------------------------------------------- report

def report(seats, res: MatchResult, bg: Background, source: str, allowed_missing,
           wrote: str | None) -> str:
    L = []
    L.append(f"Juror numbers from: {source}")
    L.append(f"Seated jurors: {len(seats)}   Background respondents: {len(bg.respondents)}")
    L.append("")
    L.append(f"{'J#':>3}  {'Sign-in name':<34} {'SurveyMonkey name':<30} Matched by")
    by_num = {s.num: s for s in seats}
    for num in sorted(by_num):
        s = by_num[num]
        if num in res.pairs:
            r, t = res.pairs[num]
            flag = "" if t in (1, 2, 9) else "   <- check"
            L.append(f"{num:>3}  {s.name[:34]:<34} {r.name[:30]:<30} {TIER_LABEL[t]}{flag}")
        elif num in res.ambiguous:
            L.append(f"{num:>3}  {s.name[:34]:<34} {'?? more than one possible':<30} NEEDS CONFIRMING")
        elif num in allowed_missing:
            L.append(f"{num:>3}  {s.name[:34]:<34} {'(no background on file)':<30} allowed")
        else:
            L.append(f"{num:>3}  {s.name[:34]:<34} {'-- not found --':<30} NEEDS CONFIRMING")

    notes = [s for s in seats if s.note]
    absent = [s for s in seats if s.attendance.strip().lower() in
              ("absent", "released", "no show", "noshow")]
    if notes or absent:
        L.append("")
        L.append("Check these - numbered, but the sign-in sheet says more:")
        for s in notes:
            L.append(f"  J{s.num} {s.name}: J# cell also says \"{s.note}\"")
        for s in absent:
            L.append(f"  J{s.num} {s.name}: Attendance says \"{s.attendance}\"")

    problems = [s for s in res.unmatched if s.num not in allowed_missing]
    if problems or res.ambiguous:
        pool = [r for r in bg.respondents
                if all(r is not p[0] for p in res.pairs.values())]
        L.append("")
        L.append("NEEDS CONFIRMING - nothing was written:")
        for s in problems:
            L.append(f"  J{s.num} {s.name}: closest unmatched background names: "
                     + ", ".join(suggestions(s, pool)))
        for num, cands in sorted(res.ambiguous.items()):
            L.append(f"  J{num} {by_num[num].name}: could be "
                     + " / ".join(c.name for c in cands))
        L.append("  Confirm with --map \"<J#>=<SurveyMonkey name>\", or --allow-missing <J#> "
                 "for a juror who has no background.")

    matched_rows = {id(p[0]) for p in res.pairs.values()}
    dropped = [r for r in bg.respondents
               if id(r) not in matched_rows and r not in res.duplicates]
    if dropped:
        L.append("")
        L.append(f"Removed - not seated ({len(dropped)}): " + ", ".join(r.name for r in dropped))
    if res.duplicates:
        L.append(f"Removed - duplicate submission ({len(res.duplicates)}): "
                 + ", ".join(f"{r.name} (row {r.row})" for r in res.duplicates))
    if bg.skipped:
        L.append(f"Removed - blank/partial rows: {len(bg.skipped)}")
    if wrote:
        L.append("")
        L.append(f"Wrote {len(res.pairs)} numbered jurors -> {wrote}")
    return "\n".join(L)


# ---------------------------------------------------------------- main

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--background", required=True, help="SurveyMonkey background export (.xlsx)")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--signin", help="recruiter sign-in workbook holding the J# column")
    src.add_argument("--roster", help="text file, one '<J#> <name>' per line")
    ap.add_argument("--sheet", help="sign-in sheet to read (default: auto-detect)")
    ap.add_argument("--out", required=True, help="where to save the numbered workbook")
    ap.add_argument("--overwrite", action="store_true", help="allow --out to be the input file")
    ap.add_argument("--map", action="append", default=[], metavar="J#=NAME",
                    help="confirmed pairing: juror number = SurveyMonkey name or Respondent ID")
    ap.add_argument("--allow-missing", default="", metavar="J#,J#",
                    help="seated jurors known to have no background submission")
    ap.add_argument("--dry-run", action="store_true", help="report only; write nothing")
    a = ap.parse_args(argv)

    if Path(a.out).resolve() == Path(a.background).resolve() and not a.overwrite:
        raise SystemExit("--out is the input file; pass --overwrite if you mean it.")

    if a.signin:
        seats, where = read_signin(a.signin, a.sheet)
        source = f"{Path(a.signin).name} [{where}]"
    else:
        seats, source = read_roster(a.roster), Path(a.roster).name
    nums = [s.num for s in seats]
    dup = sorted({n for n in nums if nums.count(n) > 1})
    if dup:
        raise SystemExit(f"Juror number(s) used twice on the sign-in list: {dup}")

    manual = {}
    for m in a.map:
        if "=" not in m:
            raise SystemExit(f"--map expects J#=NAME, got {m!r}")
        k, v = m.split("=", 1)
        manual[int(k.strip().lstrip("Jj#"))] = v
    allowed = {int(x) for x in re.findall(r"\d+", a.allow_missing)}

    bg = load_background(a.background)
    res = match(seats, bg.respondents, manual)
    blocked = bool(res.ambiguous) or any(s.num not in allowed for s in res.unmatched)
    wrote = None
    if not blocked and not a.dry_run:
        write_numbered(bg, res.pairs, a.out)
        wrote = a.out
    print(report(seats, res, bg, source, allowed, wrote))
    return 2 if blocked else 0


if __name__ == "__main__":
    sys.exit(main())
