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

# Unpaid break: a session longer than 3 hours has 30 minutes deducted, so a
# 9:00-2:30 session counts as 5 teaching hours rather than 5.5.
BREAK_THRESHOLD_HOURS = 3.0
BREAK_DEDUCTION_HOURS = 0.5

# Percentage-of-load bands.
ON_TRACK_MIN_PCT = 90.0
ON_TRACK_MAX_PCT = 110.0

STATUS_ON_TRACK = "ON TRACK"
STATUS_UNDER = "UNDER"
STATUS_OVER = "OVER"


def apply_break(gross_hours):
    """Net teaching hours for one session after the unpaid break deduction.

    A session of MORE than 3 hours loses 30 minutes:
        5.5 -> 5.0      (9:00-2:30)
        6.5 -> 6.0      (9:00-3:30)
        3.0 -> 3.0      (no deduction: not more than 3 hours)
    """
    if gross_hours is None:
        return None
    if gross_hours > BREAK_THRESHOLD_HOURS:
        return round(gross_hours - BREAK_DEDUCTION_HOURS, 2)
    return round(gross_hours, 2)


def fraction_for(teacher):
    """Teaching-load fraction, or None if the teacher is not in the table."""
    return TEACHER_FRACTIONS.get(teacher)


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
