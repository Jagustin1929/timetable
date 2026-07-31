"""
Workload model (authoritative).
===============================
Teaching-load fractions, expected hours, the unpaid-break deduction and the
on-track/under/over status thresholds.

Everything here is policy, not parsing, so it lives in one place and can be
updated each semester without touching the extraction logic.
"""

# A full teaching load is 360 delivery hours per semester.
FULLTIME_SEMESTER_HOURS = 360.0

# Teaching-load fraction per teacher (1.0 = full time).
TEACHER_FRACTIONS = {
    "Jennie Agustin":  0.5,
    "Shanna Roper":    0.5,
    "Anu Joshi":       0.5,
    "Shaun Stummer":   0.9,
    "Judi Stievenard": 0.3,
    "Martina Clark":   0.2,
    "Cathy Shay":      0.2,
    "Rima Andrews":    0.2,
    "Graham Barber":   1.0,
    "Narelle Bell":    1.0,
}

# Visiting teachers. They deliver sessions and their hours are still counted and
# shown, but they carry NO teaching load, so no fraction, expected hours,
# percentage or on-track status applies to them.
#
# This is deliberately separate from "not in TEACHER_FRACTIONS": a teacher who is
# simply missing from the table is a data gap and must be reported, whereas a
# visiting teacher having no fraction is the correct answer. Conflating the two
# would either nag about every visitor or hide a genuine omission.
VISITING_TEACHERS = {
    "Mark Kooper",
}


def is_visiting(teacher):
    """True when the teacher carries no teaching load by design."""
    return teacher in VISITING_TEACHERS

# Unpaid break: a session longer than 4 hours has 30 minutes deducted, so a
# 9:00-2:30 session counts as 5 teaching hours rather than 5.5.
#
# The threshold is 4 hours, not 3: evening classes run 5:30-9:30pm straight
# through with no break, and 4 hours must therefore count in full.
BREAK_THRESHOLD_HOURS = 4.0
BREAK_DEDUCTION_HOURS = 0.5

# Percentage-of-load bands.
ON_TRACK_MIN_PCT = 90.0
ON_TRACK_MAX_PCT = 110.0

STATUS_ON_TRACK = "ON TRACK"
STATUS_UNDER = "UNDER"
STATUS_OVER = "OVER"


def apply_break(gross_hours):
    """Net teaching hours for one session after the unpaid break deduction.

    A session of MORE than 4 hours loses 30 minutes:
        6.5 -> 6.0      (9:00-3:30)
        5.5 -> 5.0      (9:00-2:30)
        4.0 -> 4.0      (5:30-9:30pm evening class: no break, counts in full)
        3.0 -> 3.0      (no deduction)
    """
    if gross_hours is None:
        return None
    if gross_hours > BREAK_THRESHOLD_HOURS:
        return round(gross_hours - BREAK_DEDUCTION_HOURS, 2)
    return round(gross_hours, 2)


def fraction_for(teacher):
    """Teaching-load fraction, or None if the teacher carries no load.

    None covers two different cases; use is_visiting() to tell them apart. A
    visiting teacher legitimately has no fraction; anyone else with None is a
    missing entry that must be reported rather than silently treated as zero.
    """
    return TEACHER_FRACTIONS.get(teacher)


def load_label(teacher):
    """What to show in the 'Load fraction' column when there is no fraction."""
    frac = fraction_for(teacher)
    if frac is not None:
        return frac
    return "visiting" if is_visiting(teacher) else ""


def expected_hours(teacher):
    """Expected semester delivery hours = 360 x fraction. None if unknown."""
    frac = fraction_for(teacher)
    return None if frac is None else round(FULLTIME_SEMESTER_HOURS * frac, 2)


def load_percent(teacher, actual_hours):
    """Actual hours as a percentage of the teacher's expected hours."""
    exp = expected_hours(teacher)
    if not exp:
        return None
    return round((actual_hours or 0) / exp * 100.0, 1)


def status_for(percent):
    """ON TRACK 90-110%, UNDER <90%, OVER >110%."""
    if percent is None:
        return ""
    if percent < ON_TRACK_MIN_PCT:
        return STATUS_UNDER
    if percent > ON_TRACK_MAX_PCT:
        return STATUS_OVER
    return STATUS_ON_TRACK
