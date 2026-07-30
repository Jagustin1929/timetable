#!/usr/bin/env python3
"""
Agent 1 - Timetable Normaliser
==============================
Converts class-timetable Word (.docx) documents into a normalised, structured
dataset (CSV always; Excel workbook if `openpyxl` is installed).

SCOPE (deliberately limited, per project definition):
  * Structural normalisation ONLY.
  * NO co-teaching split, NO teacher renaming, NO aggregation, NO FTE logic.
    (Those happen in later phases.)
  * Sessions with no teacher are recorded and reported (pause/confirm happens
    later); Agent 1 just flags them.

Dependencies: Python standard library only for reading .docx.
              openpyxl (optional) for the .xlsx output.

Usage:
    python3 agent1_normalise.py <folder-with-docx>  [--out OUTDIR]
"""

import sys, os, re, csv, zipfile, argparse
import xml.etree.ElementTree as ET
from collections import Counter

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ---------------------------------------------------------------------------
# Class configuration: maps each source file to the class(es) it contains.
# The combined Diploma has 3 classes (one per table, in order).
# Matching is by qualification code + a distinguishing token so the odd
# filenames (double spaces, "(1)", VOF vs F2F) still resolve.
#
# A config entry is only a STARTING POINT. A document may carry more tables than
# the entry lists - e.g. the ICT40120 Cert IV file became a "combined" file
# holding one table per course stream (Programming, AI, Data). The number of
# classes is therefore always taken from the document itself, never from this
# list; see resolve_class_defs().
#
# `stream` names the course stream an entry represents, so the qualification
# LEVEL prefix can be recovered from it ("Cert IV Programming" - "Programming"
# => "Cert IV") and reused to name the other streams identically.
# ---------------------------------------------------------------------------
CLASS_CONFIG = [
    # (match tokens (all must appear, case-insensitive), [class defs by table order])
    (["bsb50520"], [
        {"class": "Diploma Library Services - PTE Evening",     "qual": "BSB50520", "delivery": "Evening"},
        {"class": "Diploma Library Services - Face to Face",    "qual": "BSB50520", "delivery": "F2F"},
        {"class": "Diploma Library Services - Fulltime VOFF",   "qual": "BSB50520", "delivery": "VOFF"},
    ]),
    (["bsb40720"],            [{"class": "Cert IV VOCF",                "qual": "BSB40720", "delivery": ""}]),
    (["ict40120"],            [{"class": "Cert IV Programming",         "qual": "ICT40120", "delivery": "",
                                "stream": "Programming"}]),
    (["ict30120", "vof"],     [{"class": "Cert III General (VOFF)",     "qual": "ICT30120", "delivery": "VOFF"}]),
    (["ict30120", "f2f"],     [{"class": "Cert III General (F2F)",      "qual": "ICT30120", "delivery": "F2F"}]),
]


def config_for(fname):
    """The CLASS_CONFIG entry for a filename, or None.

    Single place that resolves a document to its configured classes, so callers
    (the CLI and the webapp) cannot drift apart in how they match.
    """
    low = os.path.basename(fname).lower()
    return next((defs for toks, defs in CLASS_CONFIG if all(t in low for t in toks)), None)

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
UNIT_CODE_RE = re.compile(r"\b([A-Z]{3}[A-Z]{0,4}\d{3})\b")     # e.g. BSBINS516, ICTPRG440, TAEDEL301
SEM_RE = re.compile(r"Sem(?:ester)?\s*([12])\s*(20\d{2})", re.I)
QUAL_RE = re.compile(r"\b([A-Z]{2,4}\d{4,6})\b")                # e.g. ICT40120, BSB50520

# ---------------------------------------------------------------------------
# Document-level semester inference.
#
# Most source documents do NOT state their semester anywhere in the table (only
# the combined Diploma does, in its Day cells). The semester therefore has to be
# inferred from the filename / headings, e.g.:
#     "BSB40720 Cert IV VOCF FTS1 2027"  -> S1 2027   (FTS1 = Fulltime Sem 1)
#     "... Semester 2 2027 ..."           -> S2 2027
# When nothing can be determined we record it as UNKNOWN and let the extraction
# stage place the rows in whichever semester is being built. We must NEVER
# silently stamp a hardcoded semester - doing so made whole files disappear
# from any semester other than the hardcoded one.
# ---------------------------------------------------------------------------
DOC_SEM_PATTERNS = [
    re.compile(r"Sem(?:ester)?\s*([12])\s*(20\d{2})", re.I),        # Semester 1 2027
    re.compile(r"\bF\s*T\s*S\s*([12])\s*(20\d{2})\b", re.I),        # FTS1 2027
    re.compile(r"\bS\s*([12])\s*(20\d{2})\b", re.I),                # S1 2027
    re.compile(r"Sem(?:ester)?\s*([12])\s*'?(\d{2})\b", re.I),      # Semester 1 27
    re.compile(r"\bF\s*T\s*S\s*([12])\s*'?(\d{2})\b", re.I),        # FTS1 27
]


def _norm_year(y):
    """'2027' -> 2027;  '27' -> 2027."""
    y = int(y)
    return y if y >= 100 else 2000 + y


# ---------------------------------------------------------------------------
# Semester from the actual dates in the "Date of Study" column.
#
# This is the AUTHORITATIVE source: the real dates determine which semester and
# year a session belongs to. Filenames and headings are only fallbacks for rows
# that carry no usable date.
#
# Semester boundary: months 1-6 => Semester 1, months 7-12 => Semester 2.
# Validated against every row in the source set that states its own semester
# (47 rows agree, 0 disagree).
# ---------------------------------------------------------------------------
# A four-digit year may legitimately be followed by another digit, because two
# dates sometimes run together with no separator ("27/10/20271/12/2027").
# A two-digit year followed by a digit is instead a TRUNCATED four-digit year
# ("03/12/202" for 2026) and must be rejected rather than read as 2020.
DATE_RE = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(?:(\d{4})|(\d{2})(?!\d))")
TRUNCATED_DATE_RE = re.compile(r"\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{3}(?!\d)")


def semester_from_dates(text):
    """Derive (label, basis, all_labels) from real dates, e.g. 'S1 2027'.

    Handles the formats seen in the source documents:
        "22/07/2026 02/12/2026"          two dates, space separated
        "20/7/2026-30/11/2026"           range with hyphen, single-digit d/m
        "27/10/20271/12/2027"            missing separator between two dates
        "9/2/2028- 5/4/2028 24/4/2028-21/6/2028"
    The earliest valid date decides the semester. all_labels reports every
    distinct semester the dates touch so a row spanning a boundary can be
    flagged rather than silently mis-filed.
    """
    found = []
    for d, m, y4, y2 in DATE_RE.findall(text or ""):
        day, month, year = int(d), int(m), _norm_year(y4 or y2)
        if not (1 <= month <= 12 and 1 <= day <= 31):
            continue
        found.append(((year, month, day), f"S{1 if month <= 6 else 2} {year}"))
    if not found:
        return "", "", []
    found.sort()
    labels = sorted({lab for _k, lab in found})
    return found[0][1], f"dates: '{(text or '').strip()[:60]}'", labels


YEAR4_RE = re.compile(r"\b(20\d{2})\b")                      # 2027
# A bare two-digit year is only trusted in the FILENAME (e.g. "... VOF OUR 27")
# and only within a plausible range. Numbers in body text - room numbers, week
# counts, "12 months" - must never be read as a year, because a wrong year
# hint silently EXCLUDES sessions from the build.
YEAR2_RE = re.compile(r"(?:^|[\s_'\-])'?([23]\d)(?=$|[\s_.)\-])")
TRUSTED_2DIGIT_YEARS = set(range(2025, 2041))


def infer_doc_year(fname, paras=()):
    """Best-effort document-level YEAR when no full semester can be determined.

    Filenames such as "ICT30120 ... VOF OUR 27" pin the year but not the
    semester number. Recording the year lets the extraction stage keep those
    undated rows out of a build for a *different* year, while still including
    them in any semester of their own year.

    Deliberately conservative: returning a wrong year removes sessions from the
    output, so we would rather return nothing than guess.

    Returns (year_str, basis) or ('', '').
    """
    stem = os.path.splitext(os.path.basename(fname))[0]
    clean_stem = QUAL_RE.sub(" ", stem.upper())

    m = YEAR4_RE.search(clean_stem)
    if m:
        return m.group(1), f"filename: '{stem}'"
    for m in YEAR2_RE.finditer(clean_stem):
        year = _norm_year(m.group(1))
        if year in TRUSTED_2DIGIT_YEARS:
            return str(year), f"filename: '{stem}'"

    # Headings: four-digit years only.
    for p in paras:
        if not p or not p.strip():
            continue
        m = YEAR4_RE.search(QUAL_RE.sub(" ", p.upper()))
        if m:
            return m.group(1), f"heading: '{p.strip()}'"
    return "", ""


def infer_doc_semester(fname, paras=()):
    """Best-effort document-level semester label, e.g. 'S1 2027'.

    Returns (label, basis) where basis explains where it came from, or
    ('', '') when the document gives no usable semester at all.
    Qualification codes (ICT40120, BSB50520) are stripped first so their
    digits are never mistaken for a year or semester number.
    """
    stem = os.path.splitext(os.path.basename(fname))[0]
    candidates = [("filename", stem)] + [("heading", p) for p in paras if p and p.strip()]
    for basis, text in candidates:
        clean = QUAL_RE.sub(" ", text.upper())          # drop ICT40120 / BSB50520 etc.
        for pat in DOC_SEM_PATTERNS:
            m = pat.search(clean)
            if m:
                return f"S{m.group(1)} {_norm_year(m.group(2))}", f"{basis}: '{text.strip()}'"
    return "", ""


# ---------------------------------------------------------------------------
# Column mapping.
#
# Header wording varies between documents ("Day", "Day and Room",
# "Day and Channel"; "Teacher" vs "Teachers"). Locate each column by its
# header text and fall back to the historical fixed positions only for
# columns we cannot identify. Relying on position alone silently dropped the
# Teacher column whenever a document carried an extra/reordered column.
# ---------------------------------------------------------------------------
COLUMN_ALIASES = [
    ("day",      ["day"]),
    ("time",     ["time"]),
    ("unit",     ["unit of competency", "unit", "units", "competency"]),
    ("weeks",    ["week of study", "weeks", "week"]),
    ("dates",    ["date of study", "dates", "date"]),
    ("teachers", ["teachers", "teacher", "trainers", "trainer", "staff", "lecturer"]),
]
DEFAULT_POSITIONS = {"day": 0, "time": 1, "unit": 2, "weeks": 3, "dates": 4, "teachers": 5}


def map_columns(header):
    """Return (role->index, list_of_roles_that_fell_back_to_position).

    `header` is a list of header-cell strings. Matching is case-insensitive
    substring matching against COLUMN_ALIASES, longest alias first, and each
    column index is claimed by at most one role.
    """
    norm = [re.sub(r"\s+", " ", (h or "")).strip().lower() for h in header]
    colmap, taken = {}, set()
    for role, aliases in COLUMN_ALIASES:
        for alias in sorted(aliases, key=len, reverse=True):
            hit = next((i for i, h in enumerate(norm)
                        if i not in taken and h and alias in h), None)
            if hit is not None:
                colmap[role] = hit
                taken.add(hit)
                break
    fell_back = []
    for role, pos in DEFAULT_POSITIONS.items():
        if role not in colmap:
            if pos not in taken:
                colmap[role] = pos
                taken.add(pos)
            fell_back.append(role)
    return colmap, fell_back


# ---------------------------------------------------------------------------
# .docx reading (stdlib only), preserving line breaks within cells.
# ---------------------------------------------------------------------------
def para_lines(p):
    """Return the text of a <w:p> as a list of lines (splitting on <w:br>/<w:cr>)."""
    lines, cur = [], []
    for node in p.iter():
        tag = node.tag
        if tag == W + "t":
            cur.append(node.text or "")
        elif tag in (W + "br", W + "cr"):
            lines.append("".join(cur)); cur = []
        elif tag == W + "tab":
            cur.append("\t")
    lines.append("".join(cur))
    return [ln.strip() for ln in lines if ln.strip()]


def cell_lines(tc):
    out = []
    for p in tc.findall(W + "p"):
        out.extend(para_lines(p))
    return out


def read_body(path):
    """Yield ('para', text, style) and ('table', rows) in document order.
    rows = list of rows; each row = list of cells; each cell = list of lines."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    for child in body:
        if child.tag == W + "p":
            txt = " ".join(para_lines(child))
            ppr = child.find(W + "pPr")
            style = None
            if ppr is not None:
                st = ppr.find(W + "pStyle")
                style = st.get(W + "val") if st is not None else None
            yield ("para", txt, style)
        elif child.tag == W + "tbl":
            rows = []
            for tr in child.findall(W + "tr"):
                rows.append([cell_lines(tc) for tc in tr.findall(W + "tc")])
            yield ("table", rows, None)


# ---------------------------------------------------------------------------
# Class detection (used when no explicit CLASS_CONFIG entry matches a file)
# ---------------------------------------------------------------------------
def read_tables_with_headings(path):
    """Return [(rows, heading)] where `heading` is the nearest non-empty paragraph
    preceding each table. One table = one class."""
    tables, pending = [], ""
    for kind, payload, _style in read_body(path):
        if kind == "para":
            if payload.strip():
                pending = payload.strip()
        else:  # table
            tables.append((payload, pending))
            pending = ""          # a heading introduces the table right after it
    return tables


def file_mode(fname):
    """Delivery mode stated by a FILENAME, or '' when it names none."""
    low = os.path.basename(fname).lower()
    if "f2f" in low or "face to face" in low:
        return "F2F"
    if "voff" in low or re.search(r"\bvof\b", low) or "online" in low:
        return "VOFF"
    if "evening" in low or "evng" in low:
        return "Evening"
    return ""


def auto_class_defs(fname, tables):
    """Derive class definitions from the filename + per-table headings when no
    explicit CLASS_CONFIG entry matches, so any new document 'just works'."""
    stem = os.path.splitext(fname)[0]
    qm = QUAL_RE.search(stem.upper())
    qual = qm.group(1) if qm else ""
    mode = file_mode(fname)
    multi = len(tables) > 1
    defs = []
    for i, (_rows, heading) in enumerate(tables):
        name = heading or (qual or stem)
        if multi and not heading:
            name = f"{qual or stem} - Class {i + 1}"
        if mode and mode.lower() not in name.lower():
            name = f"{name} ({mode})"
        defs.append({"class": name[:80].strip(), "qual": qual, "delivery": mode})
    return defs


# ---------------------------------------------------------------------------
# Course streams within one qualification.
#
# One qualification is often delivered as several parallel courses ("streams" /
# specialisations). They arrive either as one document per stream, or as a single
# "combined" document holding ONE TABLE PER STREAM, each introduced by a heading
# that labels the stream - e.g. the ICT40120 Cert IV file carries Programming,
# AI and Data.
#
# Two rules matter:
#   1. The number of classes comes from the DOCUMENT, never from CLASS_CONFIG.
#      A combined file with 3 tables must yield 3 classes; anything else buries
#      two whole courses under a class named after the file.
#   2. A stream must produce the SAME class name whether it arrived standalone or
#      inside a combined file. The webapp supersedes old uploads by class
#      identity, so "Cert IV Programming" arriving as "ICT40120 ... Programming"
#      must still be called "Cert IV Programming" inside the combined document -
#      otherwise the old file is kept and every Programming session is counted
#      twice.
#
# ADDING A NEW STREAM: add one line to STREAM_ALIASES (and optionally its unit
# prefix to STREAM_UNIT_PREFIXES). Nothing else needs to change.
# ---------------------------------------------------------------------------
STREAM_ALIASES = [
    # (canonical stream name, phrases that identify it; matched on word boundaries)
    ("Programming",     ["programming", "software development", "coding"]),
    ("AI",              ["ai", "artificial intelligence", "machine learning"]),
    ("Data",            ["data", "data analytics", "data engineering",
                         "database", "databases", "data science"]),
    ("Networking",      ["networking", "network engineering", "networks"]),
    ("Cyber Security",  ["cyber", "cyber security", "cybersecurity", "information security"]),
    ("Web Development", ["web", "web development", "web dev", "front end web"]),
    ("Systems Admin",   ["systems administration", "system administration", "sysadmin"]),
    ("Cloud",           ["cloud", "cloud computing"]),
    ("Gaming",          ["gaming", "game development", "games"]),
    ("Digital Media",   ["digital media", "multimedia"]),
    ("General",         ["general"]),
]

# Unit-code prefixes that indicate a stream. Used only as a BACKSTOP when the
# heading above a table does not name the stream. Core units shared by every
# stream (ICTICT, BSBXCS, BSBCRT, ...) are deliberately absent, and a winner
# needs a clear majority - a stream's own units always dominate its table.
STREAM_UNIT_PREFIXES = {
    "ICTPRG": "Programming",
    "ICTAII": "AI",
    "ICTDAT": "Data",
    "ICTDBS": "Data",
    "ICTNWK": "Networking",
    "ICTCYS": "Cyber Security",
    "ICTWEB": "Web Development",
    "ICTSAS": "Systems Admin",
    "ICTCLD": "Cloud",
    "ICTGAM": "Gaming",
    "ICTDMT": "Digital Media",
}

# AQF level from the first digit of a national code's numeric part
# (ICT40120 -> 4 -> Cert IV, BSB50520 -> 5 -> Diploma).
AQF_LEVELS = {"1": "Cert I", "2": "Cert II", "3": "Cert III", "4": "Cert IV",
              "5": "Diploma", "6": "Advanced Diploma"}
LEVEL_LABEL_RE = re.compile(
    r"\b(advanced\s+diploma|diploma|cert(?:ificate)?\s*(?:iv|iii|ii|i|[1-4]))\b", re.I)
ROMAN = {"1": "I", "2": "II", "3": "III", "4": "IV"}


def qualification_base_label(fname, qual=""):
    """The qualification LEVEL label, e.g. 'Cert IV', used as the prefix of every
    stream class name so the same stream is named identically wherever it lives.

    Read from the filename ("... Cert IV Information Technology combined") and,
    failing that, from the AQF level encoded in the national code.
    """
    stem = os.path.splitext(os.path.basename(fname))[0]
    m = LEVEL_LABEL_RE.search(stem)
    if m:
        raw = re.sub(r"\s+", " ", m.group(1)).strip().lower()
        if raw.startswith("advanced"):
            return "Advanced Diploma"
        if raw.startswith("diploma"):
            return "Diploma"
        lvl = raw.split()[-1].replace("certificate", "").replace("cert", "").strip()
        return f"Cert {ROMAN.get(lvl, lvl.upper())}"
    m = re.match(r"[A-Z]{2,4}(\d)", (qual or "").upper())
    if m:
        return AQF_LEVELS.get(m.group(1), "")
    return ""


def _streams_in_text(text):
    """[(stream, matched_phrase)] for every stream named in `text`, longest
    (most specific) match first. Word-boundary matched, so 'Date of Study' never
    reads as the Data stream and 'Website' never as Web."""
    low = re.sub(r"\s+", " ", (text or "")).lower()
    hits = []
    for stream, aliases in STREAM_ALIASES:
        best = ""
        for alias in aliases:
            if re.search(r"\b" + re.escape(alias) + r"\b", low) and len(alias) > len(best):
                best = alias
        if best:
            hits.append((stream, best))
    hits.sort(key=lambda h: len(h[1]), reverse=True)
    return hits


def _stream_from_units(rows):
    """Backstop: infer a table's stream from the units it actually teaches.
    Requires a clear majority so a single shared unit cannot decide it."""
    counts = Counter()
    for row in rows:
        for cell in row:
            for code in UNIT_CODE_RE.findall(" ".join(cell)):
                prefix = re.match(r"[A-Z]+", code).group(0)
                if prefix in STREAM_UNIT_PREFIXES:
                    counts[STREAM_UNIT_PREFIXES[prefix]] += 1
    if not counts:
        return "", ""
    ranked = counts.most_common()
    top, n = ranked[0]
    runner_up = ranked[1][1] if len(ranked) > 1 else 0
    if n >= 2 and n >= 2 * max(runner_up, 1):
        return top, f"unit codes ({n} {top} unit reference(s))"
    return "", ""


def detect_stream(heading, rows):
    """Identify the course stream one table belongs to.

    Returns (stream, basis, also_matched) where `also_matched` lists any other
    streams the heading mentioned, so an ambiguous heading is flagged rather
    than quietly resolved one way.

    Signals, strongest first:
      1. the heading paragraph above the table (how the source labels streams),
      2. a title row inside the table,
      3. the unit codes being taught.
    """
    hits = _streams_in_text(heading)
    if hits:
        return hits[0][0], f"heading: '{(heading or '').strip()[:60]}'", [s for s, _a in hits[1:]]

    for row in rows[:2]:                      # a merged title row above the header
        text = " ".join(" ".join(c) for c in row)
        hits = _streams_in_text(text)
        if hits and len(row) <= 2:            # a real title row, not the column headers
            return hits[0][0], f"table title row: '{text.strip()[:60]}'", [s for s, _a in hits[1:]]

    stream, basis = _stream_from_units(rows)
    return stream, basis, []


def resolve_class_defs(fname, tables, cfg):
    """Decide the class definition for EVERY table in a document.

    The document is the authority on how many classes it holds. Three shapes are
    handled:
      * the config already describes every table (single-class files, and the
        configured combined Diploma cohorts)      -> use it unchanged
      * a combined file whose tables are course STREAMS of one qualification
        (ICT40120: Programming / AI / Data)       -> name each from its heading
      * anything else                             -> fall back to auto-detection

    Returns (defs, issues).
    """
    issues = []
    qm = QUAL_RE.search(os.path.splitext(os.path.basename(fname))[0].upper())
    qual = (cfg[0].get("qual") if cfg else "") or (qm.group(1) if qm else "")

    # 1. Config matches the document exactly - nothing to work out.
    if cfg and len(cfg) == len(tables):
        return [dict(d) for d in cfg], issues

    detected = [detect_stream(heading, rows) for rows, heading in tables]
    n_named = sum(1 for s, _b, _o in detected if s)

    # 2. Not a stream document - keep the historical behaviour and say so.
    if len(tables) < 2 or n_named == 0:
        if cfg:
            issues.append({"file": fname, "level": "WARN",
                           "msg": f"Expected {len(cfg)} table(s) for the configured classes "
                                  f"but found {len(tables)}, and no course stream could be "
                                  f"identified from the table headings. Falling back to "
                                  f"auto-detected class names - check the document structure."})
        return auto_class_defs(fname, tables), issues

    # 3. A combined, multi-stream document. Prefix every class with the
    #    qualification level so a stream is named identically wherever it lives.
    base = ""
    if cfg and cfg[0].get("stream") and cfg[0].get("class"):
        base = re.sub(r"\b" + re.escape(cfg[0]["stream"]) + r"\b", "",
                      cfg[0]["class"], flags=re.I).strip(" -")
    base = base or qualification_base_label(fname, qual) or qual or "Course"

    doc_mode = file_mode(fname)
    defs, used = [], {}
    for ti, ((rows, heading), (stream, basis, also)) in enumerate(zip(tables, detected)):
        if stream:
            name = f"{base} {stream}".strip()
        elif cfg and ti < len(cfg):
            name, basis = cfg[ti]["class"], "class configuration"
        else:
            name, basis = f"{base} Course {ti + 1}".strip(), "position in document"
            issues.append({"file": fname, "level": "WARN",
                           "msg": f"Table {ti + 1} does not name a course stream in its "
                                  f"heading ({heading.strip()[:60]!r}) and its units do not "
                                  f"identify one; it was named '{name}'. Add the stream to "
                                  f"the heading, or to STREAM_ALIASES, to name it properly."})

        # Per-table delivery mode: the heading may override the filename.
        modes = norm_mode_token(heading) or ({doc_mode} if doc_mode else set())
        mode = "/".join(sorted(modes)) if modes else ""
        if mode and mode.lower() not in name.lower():
            name = f"{name} ({mode})"

        # Two tables resolving to one name would silently merge two courses.
        if name in used:
            issues.append({"file": fname, "level": "WARN",
                           "msg": f"Tables {used[name]} and {ti + 1} both resolve to the class "
                                  f"'{name}'. They are kept apart as separate classes - check "
                                  f"whether their headings really describe different courses."})
            name = f"{name} #{ti + 1}"
        used[name] = ti + 1

        if also:
            issues.append({"file": fname, "level": "WARN",
                           "msg": f"Table {ti + 1} heading ({heading.strip()[:60]!r}) mentions "
                                  f"more than one course stream ({stream}, {', '.join(also)}); "
                                  f"used '{stream}'. Check the heading if that is wrong."})

        defs.append({"class": name[:80].strip(), "qual": qual, "delivery": mode,
                     "stream": stream, "stream_basis": basis})

    issues.append({"file": fname, "level": "INFO",
                   "msg": f"Combined document: {len(defs)} course stream(s) detected - "
                          + "; ".join(f"{d['class']} (from {d['stream_basis']})" for d in defs)})
    return defs, issues


# ---------------------------------------------------------------------------
# Normalisation helpers
# ---------------------------------------------------------------------------
def norm_mode_token(text):
    """Apply the project delivery-mode rules to a piece of text.
    Returns a set of normalised mode tokens found."""
    modes = set()
    t = text.lower()
    if "online" in t:      modes.add("VOFF")   # Online -> VOFF
    if re.search(r"\bvoff?\b", t): modes.add("VOFF")  # VOF or VOFF -> VOFF
    if "f2f" in t or "face to face" in t: modes.add("F2F")
    if "on campus" in t or "oncampus" in t: modes.add("F2F")  # (assumption - flagged)
    return modes


def parse_day_cell(lines):
    """From the 'Day' cell lines, extract weekday, semester/year, and channel/room."""
    joined = " ".join(lines)
    day = next((d for d in WEEKDAYS if re.search(r"\b" + d + r"\b", joined, re.I)
                or joined.lower().startswith(d.lower())), "")
    # also handle 'TuesdayIT-112' / 'Tuesday1A' (no space)
    if not day:
        for d in WEEKDAYS:
            if joined.lower().startswith(d.lower()):
                day = d; break
    sem = ""
    m = SEM_RE.search(joined)
    if m:
        sem = f"S{m.group(1)} {m.group(2)}"
    # channel/room = whatever remains after removing day + semester phrase
    rest = joined
    if day:
        rest = re.sub(r"\b" + day + r"\b", "", rest, flags=re.I)
    rest = SEM_RE.sub("", rest)
    channel_room = re.sub(r"\s+", " ", rest).strip(" -")
    return day, sem, channel_room


def parse_time(cell):
    """Return (start, end) from a time cell, tolerant of -, –, and spacing."""
    txt = " ".join(cell)
    txt = txt.replace("\u2013", "-").replace("\u2014", "-")
    m = re.search(r"([0-9]{1,2}[:.]?[0-9]{0,2}\s*(?:am|pm)?)\s*-\s*([0-9]{1,2}[:.]?[0-9]{0,2}\s*(?:am|pm)?)",
                  txt, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return txt.strip(), ""


def split_units(lines):
    """Split unit-of-competency cell lines into individual units (code + name).
    Trailing delivery notes (e.g. 'Graham on Campus') are captured separately."""
    units, notes, session_type = [], [], "Teaching"
    for ln in lines:
        low = ln.lower()
        if "self-directed" in low or "self directed" in low:
            session_type = "Self-directed learning"
        if "tutorial" in low:
            session_type = "Tutorial Support"
        codes = list(UNIT_CODE_RE.finditer(ln))
        if not codes:
            notes.append(ln)          # e.g. 'Tutorial Support', delivery notes, freetext
            continue
        # a line may contain multiple codes (rare) - split at each code start
        starts = [c.start() for c in codes] + [len(ln)]
        for i, c in enumerate(codes):
            seg = ln[starts[i]:starts[i + 1]].strip()
            cm = UNIT_CODE_RE.match(seg)
            code = cm.group(1) if cm else ""
            name = seg[len(code):].lstrip(" -\u2013\u2014.").strip()
            units.append({"code": code, "name": name})
    return units, notes, session_type


def normalise_file(path, class_defs=None, assumed_semester=""):
    """Return (records, issues) for one .docx file.
    If class_defs is None, the classes are auto-detected from the document
    (one table = one class, named from the preceding heading / qualification).

    `assumed_semester` is an optional explicit fallback for documents that state
    no semester of their own. It defaults to empty: rows with no determinable
    semester are marked UNKNOWN rather than being stamped with a guess, and the
    extraction stage decides where they belong.
    """
    fname = os.path.basename(path)
    records, issues = [], []
    tables = read_tables_with_headings(path)      # list of (rows, heading)

    doc_paras = [t for k, t, _s in read_body(path) if k == "para" and t.strip()]
    doc_sem, doc_sem_basis = infer_doc_semester(fname, doc_paras)
    if doc_sem:
        issues.append({"file": fname, "level": "INFO",
                       "msg": f"Document semester inferred as {doc_sem} from {doc_sem_basis}."})
    doc_year, doc_year_basis = ("", "")
    if not doc_sem:
        doc_year, doc_year_basis = infer_doc_year(fname, doc_paras)
        if doc_year:
            issues.append({"file": fname, "level": "INFO",
                           "msg": f"No semester stated; document year inferred as {doc_year} "
                                  f"from {doc_year_basis}. Rows will be included in any "
                                  f"{doc_year} semester."})

    # The document decides how many classes it holds - a "combined" file carries
    # one table per course stream and must yield one class each.
    configured = class_defs
    class_defs, class_issues = resolve_class_defs(fname, tables, class_defs)
    issues.extend(class_issues)
    if configured is None:
        issues.append({"file": fname, "level": "INFO",
                       "msg": f"Auto-detected {len(class_defs)} class(es): "
                              + "; ".join(d["class"] for d in class_defs)})

    for ti, (rows, _heading) in enumerate(tables):
        cdef = class_defs[ti] if ti < len(class_defs) else {
            "class": f"{fname} (table {ti+1})", "qual": "", "delivery": ""}
        header = [" ".join(c).strip() for c in rows[0]] if rows else []
        colmap, fell_back = map_columns(header)
        if fell_back:
            issues.append({"file": fname, "level": "WARN",
                           "msg": f"[{cdef['class']}] could not identify column(s) "
                                  f"{', '.join(sorted(fell_back))} from the header "
                                  f"{header} - using default position(s). "
                                  f"Check the document's column layout.",
                           "class": cdef["class"], "src_table": ti + 1})
        last_day = ("", "", "")     # (day, sem, channel)
        last_sem = ""
        last_date_sem = ""          # last semester derived from real dates
        last_time_val = ("", "")
        for ri, row in enumerate(rows[1:], start=1):
            def cell(role):
                """Cell lines for a logical column, located by header not position."""
                i = colmap.get(role)
                return row[i] if i is not None and i < len(row) else []

            day_cell, time_cell, unit_cell = cell("day"), cell("time"), cell("unit")
            week_cell, date_cell, teach_cell = cell("weeks"), cell("dates"), cell("teachers")

            day, sem, channel = parse_day_cell(day_cell)
            if not day and last_day[0]:                  # inherit merged day
                day, channel = last_day[0], last_day[2] if not channel else channel
            if not sem:
                sem = last_sem or (cdef.get("delivery") and "") or ""
            if day or sem:
                last_day = (day or last_day[0], sem or last_sem, channel or last_day[2])
                last_sem = sem or last_sem
            # Semester precedence. The ACTUAL DATES in the Date of Study column
            # are authoritative - they are what determines the semester and year.
            # Everything else is a fallback for rows with no usable date.
            # NEVER a hardcoded literal - see semester_from_dates().
            date_text = " ".join(date_cell).strip()
            date_sem, date_basis, date_labels = semester_from_dates(date_text)
            bad_date = TRUNCATED_DATE_RE.search(date_text)
            if bad_date:
                issues.append({"file": fname, "level": "WARN",
                               "msg": f"[{cdef['class']}] row {ri} has a malformed date "
                                      f"'{bad_date.group(0)}' (incomplete year) which was ignored. "
                                      f"Fix it in the source document. Dates: {date_text}",
                               "class": cdef["class"], "src_table": ti + 1, "src_row": ri})
            if date_sem:
                last_date_sem = date_sem
            if date_sem:
                sem_out, sem_src = date_sem, "dates"
                if len(date_labels) > 1:
                    issues.append({"file": fname, "level": "WARN",
                                   "msg": f"[{cdef['class']}] row {ri} dates span more than one "
                                          f"semester ({', '.join(date_labels)}); filed under "
                                          f"{date_sem} from the earliest date. Dates: {date_text}",
                                   "class": cdef["class"], "src_table": ti + 1, "src_row": ri})
                if sem and sem != date_sem:
                    issues.append({"file": fname, "level": "WARN",
                                   "msg": f"[{cdef['class']}] row {ri} states '{sem}' but its "
                                          f"dates give '{date_sem}'. Using the dates. "
                                          f"Dates: {date_text}",
                                   "class": cdef["class"], "src_table": ti + 1, "src_row": ri})
            elif sem:
                sem_out, sem_src = sem, "row"
            elif last_sem:
                sem_out, sem_src = last_sem, "carried"
            elif last_date_sem:
                # No date of its own (blank/merged cell): belongs with the
                # preceding dated rows of the same table.
                sem_out, sem_src = last_date_sem, "carried-dates"
            elif doc_sem:
                sem_out, sem_src = doc_sem, "inferred"
            elif assumed_semester:
                sem_out, sem_src = assumed_semester, "assumed"
            else:
                sem_out, sem_src = "", "unknown"

            tstart, tend = parse_time(time_cell)
            if not tstart and last_time_val[0]:
                tstart, tend = last_time_val
            else:
                last_time_val = (tstart, tend)

            units, notes, stype = split_units(unit_cell)
            teachers = teach_cell[:]                      # verbatim lines (Agent 1: no split/rename)
            weeks_raw = " ".join(week_cell).strip()
            dates_raw = " ".join(date_cell).strip()

            # delivery mode: class default + anything embedded in units/teacher/notes
            modes = set()
            if cdef.get("delivery"):
                modes.add(cdef["delivery"])
            for src in notes + teachers + [" ".join(unit_cell)]:
                modes |= norm_mode_token(src)
            delivery = "/".join(sorted(modes)) if modes else ""

            # skip completely empty rows
            if not (units or notes or teachers or weeks_raw or dates_raw):
                continue

            rec = {
                "source_file": fname,
                "class": cdef["class"],
                "qualification": cdef["qual"],
                "semester": sem_out,
                "semester_source": sem_src,
                "year_hint": "" if sem_out else doc_year,
                "day": day,
                "channel_or_room": channel,
                "time_start": tstart,
                "time_end": tend,
                "session_type": stype,
                "units": " | ".join(f"{u['code']} {u['name']}".strip() for u in units),
                "unit_codes": ", ".join(u["code"] for u in units if u["code"]),
                "notes": " | ".join(notes),
                "weeks_raw": weeks_raw,
                "dates_raw": dates_raw,
                "teachers": "; ".join(teachers),
                "delivery": delivery,
                "src_table": ti + 1,
                "src_row": ri,
            }
            records.append(rec)

            # ---- validation flags ----
            if not teachers:
                issues.append({"file": fname, "level": "NEEDS-TEACHER",
                               "msg": f"[{cdef['class']}] {day} {tstart}-{tend} "
                                      f"'{rec['units'] or rec['notes']}' has NO teacher.",
                               "class": cdef["class"], "src_table": ti + 1, "src_row": ri})
            if not units and not notes:
                issues.append({"file": fname, "level": "WARN",
                               "msg": f"[{cdef['class']}] row {ri}: no unit/activity text parsed."})
    return records, issues


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
FIELDS = ["source_file", "class", "qualification", "semester", "semester_source",
          "year_hint", "day", "channel_or_room",
          "time_start", "time_end", "session_type", "units", "unit_codes", "notes",
          "weeks_raw", "dates_raw", "teachers", "delivery", "src_table", "src_row"]


def write_csv(path, records):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(records)


def _sheets_from_data(records, issues):
    """Build the list of (sheet_name, header, rows) shared by both writers."""
    sheets = [("Master", FIELDS, [[r[f] for f in FIELDS] for r in records])]
    for cls in sorted({r["class"] for r in records}):
        title = re.sub(r"[\\/*?:\[\]]", " ", cls)[:31]
        sheets.append((title, FIELDS,
                       [[r[f] for f in FIELDS] for r in records if r["class"] == cls]))
    sheets.append(("Audit", ["level", "file", "class", "message"],
                   [[i.get("level"), i.get("file"), i.get("class", ""), i.get("msg")]
                    for i in issues]))
    return sheets


def _xl_escape(v):
    s = "" if v is None else str(v)
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def _col_letter(n):
    s = ""
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s


def write_xlsx(path, records, issues):
    """Write a real .xlsx using openpyxl if available, else a stdlib zip/XML writer."""
    sheets = _sheets_from_data(records, issues)
    try:
        import openpyxl
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        for name, header, rows in sheets:
            ws = wb.create_sheet(name[:31]); ws.append(header)
            for row in rows:
                ws.append(row)
        wb.save(path)
        return "openpyxl"
    except ImportError:
        pass

    # ---- stdlib fallback: xlsx is a zip of XML parts, strings written inline ----
    def sheet_xml(header, rows):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
        for ri, row in enumerate([header] + rows, start=1):
            out.append(f'<row r="{ri}">')
            for ci, val in enumerate(row):
                ref = f"{_col_letter(ci)}{ri}"
                out.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">'
                           f'{_xl_escape(val)}</t></is></c>')
            out.append('</row>')
        out.append('</sheetData></worksheet>')
        return "".join(out)

    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(len(sheets)):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" '
                  f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append('</Types>')

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')

    wb_sheets, wb_rels = [], []
    for i, (name, _, _) in enumerate(sheets):
        sid = i + 1
        wb_sheets.append(f'<sheet name="{_xl_escape(name[:31])}" sheetId="{sid}" r:id="rId{sid}"/>')
        wb_rels.append(f'<Relationship Id="rId{sid}" '
                       f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                       f'Target="worksheets/sheet{sid}.xml"/>')
    workbook_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    '<sheets>' + "".join(wb_sheets) + '</sheets></workbook>')
    workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     + "".join(wb_rels) + '</Relationships>')

    import zipfile as _zip
    with _zip.ZipFile(path, "w", _zip.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        for i, (_, header, rows) in enumerate(sheets):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", sheet_xml(header, rows))
    return "stdlib"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--out", default="output")
    ap.add_argument("--assume-semester", default="",
                    help="Semester to assume for documents that state none "
                         "(e.g. 'S1 2027'). Left empty, such rows are marked "
                         "UNKNOWN and included in whichever semester is built.")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    docs = [os.path.join(args.folder, f) for f in sorted(os.listdir(args.folder))
            if f.lower().endswith(".docx") and not f.startswith("~$")]

    all_records, all_issues = [], []
    for path in docs:
        fname = os.path.basename(path)
        cfg = config_for(fname)
        recs, iss = normalise_file(path, cfg, args.assume_semester)
        all_records.extend(recs); all_issues.extend(iss)
        tag = "" if cfg else "  (auto-detected classes)"
        sems = sorted({r["semester"] or "UNKNOWN" for r in recs})
        classes = sorted({r["class"] for r in recs})
        print(f"  {fname}: {len(recs)} session rows, {len(iss)} issue(s){tag}")
        print(f"      class(es):   {', '.join(classes) if classes else '-'}")
        print(f"      semester(s): {', '.join(sems) if sems else '-'}")

    write_csv(os.path.join(args.out, "normalised_master.csv"), all_records)
    xlsx = write_xlsx(os.path.join(args.out, "normalised.xlsx"), all_records, all_issues)

    # console report
    print("\n" + "=" * 70)
    print(f"TOTAL: {len(all_records)} normalised session rows across "
          f"{len({r['class'] for r in all_records})} classes")
    unknown_sem = [r for r in all_records if not r["semester"]]
    if unknown_sem:
        files = sorted({r["source_file"] for r in unknown_sem})
        print(f"\nROWS WITH NO STATED SEMESTER: {len(unknown_sem)} "
              f"(across {len(files)} file(s))")
        for f in files:
            n = len([r for r in unknown_sem if r["source_file"] == f])
            print(f"   - {f}: {n} row(s)")
        print("   These are included in whichever semester you build "
              "(use --assume-semester to pin them).")
        all_issues.append({"file": ", ".join(files), "level": "SEMESTER-UNKNOWN",
                           "msg": f"{len(unknown_sem)} row(s) state no semester; they will be "
                                  f"included in whichever semester is extracted."})

    needs = [i for i in all_issues if i["level"] == "NEEDS-TEACHER"]
    print(f"\nSESSIONS WITH NO TEACHER (need confirmation): {len(needs)}")
    for i in needs:
        print("   -", i["msg"])
    warns = [i for i in all_issues if i["level"] in ("WARN", "ERROR")]
    if warns:
        print(f"\nOTHER WARNINGS: {len(warns)}")
        for i in warns:
            print("   -", i["level"], i["msg"])
    print("\nWrote:", os.path.join(args.out, "normalised_master.csv"))
    print("Wrote:", os.path.join(args.out, "normalised.xlsx"), f"(via {xlsx})")


if __name__ == "__main__":
    main()
