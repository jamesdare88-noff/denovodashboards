---
name: number-jurors
description: Put the morning check-in juror numbers into a de novo mock trial's SurveyMonkey "Participant Background Questionnaire ####.xlsx" - reads the J# column from the recruiter's sign-in workbook ("Amys Tab"), matches each seated juror to their background row by name, renames rq_flag to "Juror Number", sorts the rows by juror number, deletes everyone who wasn't seated, and combines the three-words answers into Word #1 ("Helpful, Caring, and Strong"); then fills the case's "#### Juror Profiles.xlsm" (Demographic Data B-P and U, shortened wording, unused cards removed). Use whenever James asks to "number the jurors", "number the background", "add juror numbers to the SurveyMonkey background", "juror numbers for ####", or says check-in is done and the background needs numbering. Trigger on a four-digit de novo case number plus any numbering / juror number / background phrasing. Run it before demographics, attitudes, group rosters or anything else keyed on juror number. Not for Critical Insights surveys (jurors type their own number there) or for renaming photos.
---

# Number Jurors

Every downstream de novo tool keys on juror number - demographics, attitudes,
group rosters, deliberation groups. The SurveyMonkey background export has no
juror number; it is keyed by name. At morning check-in James numbers the
jurors and the numbers land in the recruiter's sign-in workbook. This skill
carries them across, which James used to do by hand.

## Inputs

1. **The background export** - `Participant Background Questionnaire ####.xlsx`,
   downloaded into `JAMES DASHBOARDS\#### <caption>\Survey Monkey Backgrounds\`
   (the folder is sometimes spelled `SurveyMonkey Backgrounds`). The untouched
   original also lives in the case's `Consulting\Project Docs\Results\`.
2. **The recruiter sign-in workbook** - in the case's `AdminPM\Project Management\`
   folder. Nelson sheets are named `Shared <City>, <ST> <Day> <Date> (...).xlsx`;
   Nationwide sheets are `####_De_novo_<Place>_<date>.xlsx`. Both carry
   **Amy's tab** (`Amys Tab`): `Attendance | J# | Rec Names | Match/No Match | ... | SM Names`.

**Juror numbers always come from the `J#` column of Amy's tab. Nothing else.**
Never take them from the `#### Juror Profiles.xlsm` workbook, the recruiter's
main-sheet `#` column (that's the recruit number), photo filenames, or row
order. The script reads Amy's tab only and stops if the workbook doesn't have one.

If either file is missing, ask. If Amy's tab has no `J#` numbers yet, check-in
isn't finished, so say so and wait.

The Microsoft 365 connector reads `.xlsx` as text but cannot hand back the file
itself, so get real copies: stage them from James's device (`device_stage_files`),
or ask him to attach them.

## Run

`S` is this skill's `scripts/` directory.

```bash
python3 $S/number_jurors.py \
  --background "Participant Background Questionnaire 1425.xlsx" \
  --signin "Shared Philadelphia, PA Thursday October 1st (7_15 AM - 6_30 PM).xlsx" \
  --out "/tmp/numbered/Participant Background Questionnaire 1425.xlsx"
```

Needs `openpyxl`. The script:

- reads the `J#` column of Amy's tab (any spelling: `Amys Tab`, `Amy's tab`);
  a J# cell like `21 MIA-CUT` counts as juror 21 and the note is reported;
  a cell with no number (`GHOSTED`, blank) means not seated;
- reads `Rec Names` even when the formulas were never calculated;
- matches names after stripping `***EB`, `late`, `- released at lunch...`,
  `Jr.`, middle names and initials, quoted nicknames and capitals; hyphenated
  surnames match either half;
- renames a SurveyMonkey column to **`Juror Number`** and writes the numbers
  there - `rq_flag` normally, or an existing `Juror Number(s)` column (1394
  keeps it in column I), or a new column after `Custom Data 1` if neither
  exists. Keep the header singular: the attitudes dashboard requires exactly
  `Juror Number` and stops on `Juror Numbers`, and the readiness board looks
  for the singular too;
- sorts by juror number, moving **whole rows** so Respondent ID and timestamps
  stay with their person;
- deletes non-seated respondents, duplicate submissions (keeps the fullest),
  blank partials and `Column1...` table artefact rows;
- combines the **three words** ("What three words would your family and
  friends use to describe you?") into **Word #1**, exactly as James's
  `CleanAndConcatenateXYZ` macro does: trim, capitalize the first letter,
  lowercase the rest, join as `A, B, and C` (`Helpful, Caring, and Strong`).
  Word #2 and #3 are left as they are. A juror who gave only two words gets
  `A and B`. Running it again on an already-combined file doesn't
  double-combine. `--no-three-words` skips this step.

## It never guesses

A seated juror who can't be matched with confidence stops the run: exit code
**2**, nothing written, and the report lists the closest background names.
Two seated jurors who share a first name are never matched on first name.

Show James the `NEEDS CONFIRMING` lines and ask. Then re-run with:

- `--map "20=Shannon"` - he confirmed the pairing (exact SurveyMonkey name or
  Respondent ID);
- `--allow-missing 9` - that juror never filled out the background.

Never pass `--map` for a pairing you worked out yourself. A wrong number
spreads into every dashboard built after it, and nobody downstream can see it.

## Before saving

Read the report.

- Rows marked **`<- check`** (`first name only`, `similar spelling`,
  `close spelling`) were matched, but on less than a full name. List them for
  James in one short message (`J20 Shannon McCleary -> "Shannon"`) and save
  once he says OK. If every match is `exact` or `first + last`, save straight away.
- **`Check these`** lines are jurors with a number plus a note (`MIA-CUT`,
  `gone too long`), or with Attendance saying ABSENT or Released. They are
  numbered and kept. Mention them. If James wants a cut juror out, re-run with
  `--exclude 21` - don't edit the output by hand.
- The seated count should equal the panel size James expects (usually 24-28).

## Deliver

Save the numbered workbook over the downloaded copy in
`JAMES DASHBOARDS\#### <caption>\Survey Monkey Backgrounds\` (keep the
filename; the untouched original is still in `Consulting\Project Docs\Results\`).
Commit it with `device_commit_files` when you're working on his device, or send
it with `SendUserFile` otherwise.

Reply in two or three lines: how many jurors were numbered, how many people
were removed, and anything flagged. Don't paste the whole table unless he asks.

## Then fill the Juror Profiles

The numbered background feeds `JAMES DASHBOARDS\#### <caption>\#### Juror Profiles.xlsm`.
James used to copy background columns J-X into the workbook's
**Demographic Data** sheet (B-P) and the combined three words into column U,
then click the shortening buttons. `fill_profiles.py` does all of it:

```bash
python3 $S/fill_profiles.py \
  --background "Participant Background Questionnaire 1425.xlsx" \
  --profiles "1425 Juror Profiles.xlsm" \
  --case-title "1425 Connor v. Harrow-Imprimis" \
  --out "/tmp/numbered/1425 Juror Profiles.xlsm"
```

- Juror N goes in row N+1, because card N's text boxes read row N+1.
- Race, education, employment, relationship and income are shortened exactly
  as the workbook's `Replace...Values` macros do (`Black`, `Bachelor’s`,
  `Single`, `$75k-$99,9k` ...).
- The card text and the text boxes are refreshed, and the workbook is set to
  recalculate when Excel opens it.
- On a panel smaller than the workbook (24 of 27), cards 25-27 and their
  photos are deleted and the print area is trimmed - **only** on a tidied
  workbook (see below). On an untidied one it says so and leaves them.
- `--case-title` replaces the old case name in the print header. Use the case
  folder's name.
- Long job titles are listed: they wrap on the card, so mention them.

It never saves through openpyxl - that would drop every text box and card.
Only the Demographic Data sheet, the card drawing and workbook.xml change;
macros, photos and card formatting are copied through untouched.

After it runs, James opens the file in Excel once, adds the photos with his
photo buttons, and saves. Commit the filled workbook over the case's copy.

Photos are still his: `InsertPhotosFromFolder` down column H, then
`RenamePicturesBasedOnLocation` (names them `Juror 1`, `Juror 2` ...), then
`CenterMatchingObjects` (moves each `Juror N` photo onto the hidden `Juror N`
cell label under its card). Don't rename photos or clear those labels.

## One-time: tidy a Juror Profiles master

`tidy_profiles_master.py` gives every card the same build and names -
`Juror N` (photo), `Juror N Card` (group) holding `Juror N Background`,
`Juror Na` (profile text), `Juror Nb` (three words), `Juror N Label`,
`Juror N Line` - groups any card left as loose pieces, deletes leftovers
hidden under the cards, fixes the hidden cell labels, restores the Age Range
formulas and resets the print header to `#### Case Name`.

```bash
python3 $S/tidy_profiles_master.py --profiles "Juror Profiles Master Copy.xlsm" \
  --out "Juror Profiles Master (tidied).xlsm"
```

Run it on the master once (done for the October 2026 master) or on an old
case's workbook before filling it. It writes a new file; keep the original
until James has opened the tidied copy in Excel and checked the cards.
Running it again on a tidied workbook changes nothing.
