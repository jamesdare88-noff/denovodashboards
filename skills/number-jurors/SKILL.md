---
name: number-jurors
description: Put the morning check-in juror numbers into a de novo mock trial's SurveyMonkey "Participant Background Questionnaire ####.xlsx" - reads the J# column from the recruiter's sign-in workbook ("Amys Tab"), matches each seated juror to their background row by name, renames rq_flag to "Juror Number", sorts the rows by juror number and deletes everyone who wasn't seated. Use whenever James asks to "number the jurors", "number the background", "add juror numbers to the SurveyMonkey background", "juror numbers for ####", or says check-in is done and the background needs numbering. Trigger on a four-digit de novo case number plus any numbering / juror number / background phrasing. Run it before demographics, attitudes, group rosters or anything else keyed on juror number. Not for Critical Insights surveys (jurors type their own number there) or for renaming photos.
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
   Nationwide sheets are `####_De_novo_<Place>_<date>.xlsx`. Both carry an
   **`Amys Tab`** sheet: `Attendance | J# | Rec Names | Match/No Match | ... | SM Names`.
   `J#` is the juror number. The main sheet's `#` column is the recruit number -
   never use it.

If either file is missing, ask. If the sign-in sheet has no `J#` yet, check-in
isn't finished - say so. If James would rather type the numbers, take a list
(`1 Jane Doe` per line) and use `--roster` instead of `--signin`.

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

- finds the `J#` column on whichever sheet has one (`--sheet` to force one);
  a J# cell like `21 MIA-CUT` counts as juror 21 and the note is reported;
  a cell with no number (`GHOSTED`, blank) means not seated;
- reads `Rec Names` even when the formulas were never calculated;
- matches names after stripping `***EB`, `late`, `- released at lunch...`,
  `Jr.`, middle names and initials, quoted nicknames and capitals; hyphenated
  surnames match either half;
- writes the number into `rq_flag` renamed to **`Juror Number`**, or into an
  existing `Juror Number(s)` column (1394 keeps it in column I), or inserts the
  column after `Custom Data 1`;
- sorts by juror number, moving **whole rows** so Respondent ID and timestamps
  stay with their person;
- deletes non-seated respondents, duplicate submissions (keeps the fullest),
  blank partials and `Column1...` table artefact rows.

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
  that juror off the list (`--roster`) - don't edit the output by hand.
- The seated count should equal the panel size James expects (usually 24-28).

## Deliver

Save the numbered workbook over the downloaded copy in
`JAMES DASHBOARDS\#### <caption>\Survey Monkey Backgrounds\` (keep the
filename; the untouched original is still in `Consulting\Project Docs\Results\`).
Commit it with `device_commit_files` when you're working on his device, or send
it with `SendUserFile` otherwise.

Reply in two or three lines: how many jurors were numbered, how many people
were removed, and anything flagged. Don't paste the whole table unless he asks.
