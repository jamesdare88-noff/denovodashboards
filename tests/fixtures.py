"""Synthetic SurveyMonkey background exports for the tests (no real people)."""
import openpyxl

META = ["Respondent ID", "Collector ID", "Start Date", "End Date", "IP Address",
        "Email Address", "First Name", "Last Name", "Custom Data 1"]
DEMO = ["What is your name? (First AND Last Name)", "What is your gender?",
        "What is your current age?", "Are you of Hispanic, Latin, or Spanish origin?",
        "What is your race?", "What is the highest level of education you have completed?",
        "What is your current employment status?",
        "What is your current or most recent job title?", "In what industry do/did you work?",
        "What is your relationship status?", "Which best describes your current living situation?",
        "What is your political affiliation?",
        "Which of these describes you, politically (even if none is a perfect fit)?",
        "How much total combined money did all members of your HOUSEHOLD earn last year, before taxes?"]
SUB = ["Open-Ended Response", "Response", "Open-Ended Response"] + ["Response"] * 4 + \
      ["Open-Ended Response"] + ["Response"] * 6
WORDS_Q = "What three words would your family and friends use to describe you?"

# (name, gender, age, hispanic, race, education, employment, job, industry,
#  relationship, living, affiliation, leaning, income, (w1, w2, w3))
PEOPLE = [
    ("Ava Thompson", "Female", 29, "No", "White", "Bachelor’s degree",
     "Working part-time for employer", "Substitute teacher", "Education",
     "Single, never married", "Rent home", "Democrat", "Very liberal", "$75,000-$99,999",
     ("Helpful", "caring ", "STRONG")),
    ("Ben Carter", "Male", 42, "No", "Black or African American", "GED/High school diploma",
     "Self-employed", "Delivery", "Wholesale/Retail", "Married",
     "Own home, with mortgage or loan", "Democrat", "Somewhat liberal", "$30,000-$49,999",
     ("Hard work", "Family care", "Compassion")),
    ("Chloe Diaz", "Female", 61, "Yes", "None of these applies; I identify as Latino/Hispanic",
     "Some college, no degree", "Working full-time for employer",
     "Assistant housekeeping director at a large hotel", "Other Non-Medical or Educational Service",
     "Cohabitating with significant other, not married or civil union",
     "Living with family or friend(s)", "Independent", "Slightly conservative",
     "$200,000 or more", ("Loyal", None, "Kind")),
]


def make_numbered_background(path, people=PEOPLE, numbers=None):
    """A numbered background like number_jurors.py writes (Juror Number in J,
    three words not yet combined)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(META + ["Juror Number"] + DEMO + ["Some attitude question", WORDS_Q, None, None])
    ws.append([None] * 10 + SUB + ["Response", "Word #1", "Word #2", "Word #3"])
    for i, p in enumerate(people):
        n = numbers[i] if numbers else i + 1
        ws.append([115186000000 + i, 1, None, None, None, None, None, None, None, n]
                  + list(p[:14]) + ["Agree"] + list(p[14]))
    wb.save(path)
