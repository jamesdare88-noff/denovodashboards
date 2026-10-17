# denovodashboards

Tools behind James's de novo Jury Consulting dashboards.

## Number Jurors (`skills/number-jurors`)

Puts the morning check-in juror numbers into the SurveyMonkey
`Participant Background Questionnaire ####.xlsx`, replacing the manual step.

- Reads `J#` from Amy's tab in the recruiter sign-in workbook - the only
  source of juror numbers (never Juror Profiles).
- Matches each seated juror to their background row by name. Recruiter
  suffixes (`***EB`, `late`, notes), nicknames, middle names, capitals and
  hyphenated surnames are all handled.
- Renames `rq_flag` to `Juror Number` (singular, which the attitudes dashboard
  requires), sorts by juror number (whole rows), and deletes everyone who
  wasn't seated. `--exclude J#` leaves out a juror who was cut.
- Combines the three-words answers into Word #1 the way the
  `CleanAndConcatenateXYZ` macro does (`Helpful, Caring, and Strong`).
- Stops and asks rather than guessing when a name can't be matched.

As a Claude skill: upload `skills/number-jurors/` and ask Claude to
"number the jurors for 1425".

From the command line:

```bash
pip install -r requirements.txt
python skills/number-jurors/scripts/number_jurors.py \
  --background "Participant Background Questionnaire 1425.xlsx" \
  --signin "Shared Philadelphia, PA Thursday October 1st (7_15 AM - 6_30 PM).xlsx" \
  --out "numbered/Participant Background Questionnaire 1425.xlsx"
```

Exit code 2 means a juror needs confirming. The report names the closest
matches; re-run with `--map "J#=SurveyMonkey name"` or `--allow-missing J#`.

Tests: `python -m pytest tests`
