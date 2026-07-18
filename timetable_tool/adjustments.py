"""
Manual scheduling adjustments (audit trail).
=============================================
Each adjustment overrides what a source document says, to resolve real-world
clashes the source can't express. Every one is applied AND logged in the
workbook's "Adjustments & audit" sheet so a reviewer can see what changed and why.

NOTE (source corrected): the Cert III F2F/VOF documents were later updated at the
source to (a) assign the previously-missing teachers, (b) clean the duplicated
name on the Thursday networking session, and (c) split the Thursday Cloud row into
ICTCLD301 (Graham, wks 10-14) and ICTICT310 (Anu, wks 14-18). Because of that the
earlier ANU-1 / ANU-3 / GRAHAM-1 adjustments are no longer needed and were removed
to avoid double-counting. Only the corrections NOT covered by the source remain.

Matching a session (all provided keys must match):
    class_contains : substring of the class name  (e.g. "Cert IV Programming")
    day            : exact weekday
    time_start     : prefix of the start time     (e.g. "9:00")
    unit           : substring of the units/activity text (e.g. "ICTPRG302")

Operations:
    remove_teacher    - drop a teacher from the matched session
    set_teacher_weeks - restrict/relocate a teacher's weeks on the session
    add_teacher       - add a teacher (optionally only certain weeks / mode)
"""

ADJUSTMENTS = [
    # Anu stops the Thursday ICT40120 session at wk12 so she can pick up the Cert III
    # VOFF Thursday unit (ICTICT310, wks 14-18) which the source now assigns to her.
    # (ICT40120 is a separate document and still shows her for all 18 weeks.)
    {
        "id": "ANU-2", "op": "set_teacher_weeks", "teacher": "Anu Joshi", "weeks": "1-12",
        "match": {"class_contains": "Cert IV Programming", "day": "Thursday",
                  "time_start": "9:00", "unit": "ICTPRG302"},
        "reason": "Anu stops the Thursday ICT40120 session after wk12 (8 Nov 2026); "
                  "Shaun Stummer continues wks 1-18. Prevents a clash with her Cert III "
                  "VOFF Thursday unit (wks 14-18).",
    },

    # Shanna can't be in two places in wk9 (she moves to the VOFF Wednesday session),
    # so her F2F ICTWEB304 becomes wks 1-8 and Jennie covers the F2F wk9.
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
# cohort documents. Merged (counted once) when: same teacher, day, time, session
# type; both classes are the Cert III F2F and VOFF; and the activity text says
# "combined".
MERGE_COMBINED_CERT3 = True
COMBINED_LABEL = "Cert III General (F2F+VOFF combined)"
