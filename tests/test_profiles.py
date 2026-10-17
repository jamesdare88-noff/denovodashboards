"""Tests for fill_profiles.py and tidy_profiles_master.py.

tests/data/juror_profiles_fixture.xlsm is James's Juror Profiles master with
every photo and card image replaced by a blank square - same sheets, card
shapes, links, buttons and macros, no real people.
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skills/number-jurors/scripts"))
sys.path.insert(0, str(ROOT / "tests"))
import fill_profiles as fp  # noqa: E402
import fixtures  # noqa: E402
import tidy_profiles_master as tp  # noqa: E402
from xlsx_package import NS, Package, cell_text, get_cell, shared_strings  # noqa: E402

FIXTURE = ROOT / "tests/data/juror_profiles_fixture.xlsm"


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def shape_names(path):
    pkg = Package(path)
    dr = pkg.xml(pkg.drawing_part(pkg.sheet_part("Juror Profiles")))
    return [tp.name_of(e) for e in dr.iter() if e.tag in tp.SHAPE_TAGS]


@pytest.fixture(scope="module")
def tidied(tmp_path_factory):
    out = tmp_path_factory.mktemp("t") / "tidied.xlsm"
    pkg = Package(FIXTURE)
    rep = tp.tidy(pkg)
    pkg.save(out)
    return out, rep


# ------------------------------------------------------------------ tidy

def test_tidy_master(tidied):
    out, rep = tidied
    assert rep["problems"] == []
    assert rep["cards"] == list(range(1, 28))
    assert rep["grouped"] == [10, 23, 27]
    assert len(rep["deleted"]) == 4
    names = shape_names(out)
    for n in range(1, 28):
        for need in (f"Juror {n}", f"Juror {n} Card", f"Juror {n}a", f"Juror {n}b",
                     f"Juror {n} Background", f"Juror {n} Label", f"Juror {n} Line"):
            assert names.count(need) == 1, need
    # nothing left with an Excel default name except the buttons
    leftovers = [n for n in names if re.match(r"(Picture|TextBox|Rectangle|Group|Straight)", n)]
    assert leftovers == []


def test_tidy_keeps_links_formatting_and_macros(tidied):
    out, _ = tidied
    before, after = parts(FIXTURE), parts(out)
    assert after["xl/vbaProject.bin"] == before["xl/vbaProject.bin"]
    assert all(after[n] == before[n] for n in before if n.startswith("xl/media/"))
    changed = {n for n in before if before[n] != after.get(n)}
    assert changed <= {"xl/drawings/drawing2.xml", "xl/worksheets/sheet1.xml",
                       "xl/worksheets/sheet2.xml", "xl/calcChain.xml", "xl/workbook.xml",
                       "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}
    pkg = Package(out)
    dr = pkg.xml(pkg.drawing_part(pkg.sheet_part("Juror Profiles")))
    links = {tp.name_of(sp): sp.get("textlink") for sp in dr.iter(tp.X + "sp")
             if sp.get("textlink")}
    for n in range(1, 28):
        assert links[f"Juror {n}a"] == f"'Demographic Data'!A{34 + n}"
        assert links[f"Juror {n}b"] == f"'Demographic Data'!B{34 + n}"
    # text formatting untouched: same number of Raleway runs as before
    orig = Package(FIXTURE)
    odr = orig.xml(orig.drawing_part(orig.sheet_part("Juror Profiles")))
    count = lambda d: len(d.findall(".//a:latin[@typeface='Raleway']", NS))  # noqa: E731
    assert count(dr) == count(odr)


def test_tidy_cell_labels_and_data_sheet(tidied):
    out, _ = tidied
    pkg = Package(out)
    sst = shared_strings(pkg)
    cards = pkg.xml(pkg.sheet_part("Juror Profiles"))
    for n in range(1, 28):
        band, left = (n - 1) // 2, n % 2
        r = 4 + 5 * band
        assert cell_text(cards, f"{'C' if left else 'F'}{r}", sst) == f"Juror {n}"
        assert cell_text(cards, f"{'D' if left else 'G'}{r}", sst) == f"Juror {n}a"
    assert all(cell_text(cards, f"H{r}", sst) == "" for r in range(1, 72))
    assert cell_text(cards, "F69", sst) == ""          # no 'Juror 28'
    assert "#### Case Name" in cards.findtext("m:headerFooter/m:oddHeader", namespaces=NS)
    data = pkg.xml(pkg.sheet_part("Demographic Data"))
    for r in range(2, 29):
        f = get_cell(data, f"R{r}").find("m:f", NS)
        assert f is not None and f"E{r}>=18" in f.text
    assert get_cell(data, "A63").find("m:f", NS) is None
    assert "xl/calcChain.xml" not in parts(out)


def test_tidy_twice_changes_nothing_more(tidied, tmp_path):
    out, _ = tidied
    pkg = Package(out)
    rep = tp.tidy(pkg)
    assert rep["problems"] == [] and rep["grouped"] == [] and rep["deleted"] == []
    assert rep["renamed"] == 0 and rep["labels"] == []


# ------------------------------------------------------------------ fill

def test_shorten_matches_the_macros():
    assert fp.shorten("race", "Black or African American") == "Black"
    assert fp.shorten("race", "None of these applies; I identify as Latino/Hispanic") == "Hispanic"
    assert fp.shorten("education", "Bachelor’s degree") == "Bachelor’s"
    assert fp.shorten("education", "Bachelor's degree") == "Bachelor’s"   # straight apostrophe
    assert fp.shorten("education", "Master’s degree") == "Master's"
    assert fp.shorten("marital", "Single, never married") == "Single"
    assert fp.shorten("income", "$200,000 or more") == "$200k+"
    assert fp.shorten("employment", "Retired") == "Retired"
    assert fp.shorten("job", "Black or African American") == "Black or African American"


def test_generation_lookup():
    table = [(19, "Gen Z"), (29, "Millennial"), (45, "Gen X"), (61, "Boomer"), (80, "Silent Gen")]
    assert fp.generation(19, table) == "Gen Z"
    assert fp.generation(28, table) == "Gen Z"
    assert fp.generation(29, table) == "Millennial"
    assert fp.generation(61, table) == "Boomer"
    assert fp.generation(18, table) is None
    assert fp.generation(None, table) is None


def test_fill_three_jurors(tmp_path, capsys):
    bg, out = tmp_path / "bg.xlsx", tmp_path / "out.xlsm"
    fixtures.make_numbered_background(bg)
    assert fp.main(["--background", str(bg), "--profiles", str(FIXTURE), "--out", str(out),
                    "--case-title", "9999 Test v. Example", "--keep-extra-cards"]) == 0
    pkg = Package(out)
    sst = shared_strings(pkg)
    d = pkg.xml(pkg.sheet_part("Demographic Data"))
    row = [cell_text(d, f"{c}2", sst) for c in "BCDEFGHIJKLMNOP"]
    assert row == ["1", "Ava Thompson", "Female", "29", "No", "White", "Bachelor’s",
                   "Part-time/Emp.", "Substitute teacher", "Education", "Single", "Rent home",
                   "Democrat", "Very liberal", "$75k-$99,9k"]
    assert cell_text(d, "G4", sst) == "Hispanic" and cell_text(d, "L4", sst) == "Cohabitating"
    assert cell_text(d, "U2", sst) == "Helpful, Caring, and Strong"
    assert cell_text(d, "U4", sst) == "Loyal and Kind"
    assert cell_text(d, "C5", sst) == ""                       # juror 4: no one
    assert cell_text(d, "A35", sst) == ("Age 29 - Millennial\nWhite - Female\nSubstitute teacher\n"
                                        "Bachelor’s\nSingle\nVery liberal\nDemocrat")
    assert cell_text(d, "A38", sst) == "#N/A"
    assert cell_text(d, "S3", sst) == "Millennial"
    # the card text boxes show the new text
    dr = pkg.xml(pkg.drawing_part(pkg.sheet_part("Juror Profiles")))
    shown = {sp.get("textlink"): sp.findtext(".//a:fld/a:t", namespaces=NS)
             for sp in dr.iter(tp.X + "sp") if sp.get("textlink")}
    assert shown["'Demographic Data'!B35"] == "Helpful, Caring, and Strong"
    assert shown["'Demographic Data'!A36"].startswith("Age 42 - Millennial\nBlack - Male")
    # header, recalc, and nothing else touched
    cs = pkg.xml(pkg.sheet_part("Juror Profiles"))
    assert cs.findtext("m:headerFooter/m:oddHeader", namespaces=NS).endswith(
        "&R&14 9999 Test v. Example")
    assert pkg.xml("xl/workbook.xml").find("m:calcPr", NS).get("fullCalcOnLoad") == "1"
    before, after = parts(FIXTURE), parts(out)
    assert list(before) == list(after)
    assert {n for n in before if before[n] != after[n]} == {
        "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml",
        "xl/drawings/drawing2.xml"}
    text = capsys.readouterr().out
    assert "Filled 3 jurors" in text and "Long job title" in text


def test_fill_24_removes_cards_25_to_27(tidied, tmp_path, capsys):
    master, _ = tidied
    people = []
    for i in range(24):
        p = list(fixtures.PEOPLE[i % 3])
        p[0] = f"Person {i + 1:02d}"
        people.append(tuple(p))
    bg, out = tmp_path / "bg.xlsx", tmp_path / "out.xlsm"
    fixtures.make_numbered_background(bg, people)
    assert fp.main(["--background", str(bg), "--profiles", str(master), "--out", str(out)]) == 0
    names = shape_names(out)
    assert "Juror 24 Card" in names and "Juror 24" in names
    assert not any(n in names for n in ("Juror 25 Card", "Juror 25", "Juror 27 Card", "Juror 27"))
    wb = Package(out).xml("xl/workbook.xml")
    pa = [d.text for d in wb.iterfind("m:definedNames/m:definedName", NS)
          if d.get("name") == "_xlnm.Print_Area"][0]
    assert pa.endswith("$G$61")
    assert "Deleted unused card(s) [25, 26, 27]" in capsys.readouterr().out


def test_fill_untidied_master_leaves_extra_cards(tmp_path, capsys):
    bg, out = tmp_path / "bg.xlsx", tmp_path / "out.xlsm"
    fixtures.make_numbered_background(bg)
    assert fp.main(["--background", str(bg), "--profiles", str(FIXTURE), "--out", str(out)]) == 0
    assert "Extra cards NOT deleted" in capsys.readouterr().out
    assert "Juror 27" in shape_names(out)


def test_fill_rejects_juror_beyond_cards(tmp_path):
    bg = tmp_path / "bg.xlsx"
    fixtures.make_numbered_background(bg, numbers=[1, 2, 30])
    with pytest.raises(SystemExit, match="only has 27 cards"):
        fp.main(["--background", str(bg), "--profiles", str(FIXTURE),
                 "--out", str(tmp_path / "o.xlsm")])


def test_fill_rejects_unexpected_background_layout(tmp_path):
    import openpyxl
    bg = tmp_path / "bg.xlsx"
    fixtures.make_numbered_background(bg)
    wb = openpyxl.load_workbook(bg)
    wb.active.cell(1, 14).value = "Something new"        # where race should be
    wb.save(bg)
    with pytest.raises(SystemExit, match="layout has changed"):
        fp.main(["--background", str(bg), "--profiles", str(FIXTURE),
                 "--out", str(tmp_path / "o.xlsm")])


def test_fill_refuses_to_overwrite_input(tmp_path):
    bg, prof = tmp_path / "bg.xlsx", tmp_path / "p.xlsm"
    fixtures.make_numbered_background(bg)
    shutil.copy(FIXTURE, prof)
    with pytest.raises(SystemExit):
        fp.main(["--background", str(bg), "--profiles", str(prof), "--out", str(prof)])
