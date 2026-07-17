"""
Manual scheduling adjustments (audit trail).
=============================================
Each adjustment overrides what a source document says, to resolve real-world
clashes/combined sessions. Every one is applied AND logged in the workbook's
"Adjustments applied" audit sheet so a reviewer can see exactly what changed
and why.

Matching a session (all provided keys must match):
    class_contains : substring of the class name  (e.g. "Cert IV Programming")
    day            : exact weekday
    time_start     : prefix of the start time     (e.g. "9:00")
    unit           : substring of the units/activity text (e.g. "ICTPRG302")

Operations:
    remove_teacher    - drop a teacher from the matched session
    set_teacher_weeks - restrict/relocate a teacher's weeks on the session
    add_teacher       - add a teacher (optionally only certain weeks / mode)
    (combined tutorials that appear in both Cert III cohorts are auto-merged
     separately by the "combined" rule.)
"""

ADJUSTMENTS = [
    # --- Anu Joshi: move from ICT40120 Thursday to Cert III VOFF Thursday mid-semester ---
    {
        "id": "ANU-1", "op": "remove_teacher", "teacher": "Anu Joshi",
        "match": {"class_contains": "VOFF", "day": "Thursday", "time_start": "9:00",
                  "unit": "ICTNWK311"},
        "reason": "Anu teaches Thursday VOFF only from wk14; source listed her on the "
                  "wks 1-9 networking session (with a duplicated 'Graham Barber') - removed.",
    },
    {
        "id": "ANU-2", "op": "set_teacher_weeks", "teacher": "Anu Joshi", "weeks": "1-12",
        "match": {"class_contains": "Cert IV Programming", "day": "Thursday",
                  "time_start": "9:00", "unit": "ICTPRG302"},
        "reason": "Anu stops the Thursday ICT40120 session after wk12 (8 Nov 2026). "
                  "Shaun Stummer continues wks 1-18.",
    },
    {
        "id": "ANU-3", "op": "add_teacher", "teacher": "Anu Joshi", "weeks": "14-18",
        "mode": "VOFF",
        "match": {"class_contains": "VOFF", "day": "Thursday", "time_start": "9:00",
                  "unit": "ICTCLD301"},
        "reason": "Anu takes the Thursday VOFF (Cloud) session from wk14 (15 Nov 2026). "
                  "Wks 10-13 remain unassigned - flagged.",
    },

    # --- Shanna Roper / Jennie Agustin: Wednesday wk9 web-pages ---
    {
        "id": "WED9-1", "op": "set_teacher_weeks", "teacher": "Shanna Roper", "weeks": "1-8",
        "match": {"class_contains": "F2F", "day": "Wednesday", "time_start": "9:00",
                  "unit": "ICTWEB304"},
        "reason": "Shanna moves to the VOFF Wednesday session in wk9, so her F2F "
                  "ICTWEB304 weeks become 1-8.",
    },
    {
        "id": "WED9-2", "op": "add_teacher", "teacher": "Jennie Agustin", "weeks": "9",
        "mode": "F2F",
        "match": {"class_contains": "F2F", "day": "Wednesday", "time_start": "9:00",
                  "unit": "ICTWEB304"},
        "reason": "Jennie Agustin takes the extra wk9 for F2F ICTWEB304 (Shanna is on VOFF).",
    },
]

# Combined tutorials that are physically ONE session but listed in both Cert III
# cohort documents. They are merged (counted once) when: same teacher, day, time,
# session type; both classes are the Cert III F2F and VOFF; and the activity text
# contains "combined".
MERGE_COMBINED_CERT3 = True
COMBINED_LABEL = "Cert III General (F2F+VOFF combined)"
