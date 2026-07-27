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

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ---------------------------------------------------------------------------
# Class configuration: maps each source file to the class(es) it contains.
# The combined Diploma has 3 classes (one per table, in order).
# Matching is by qualification code + a distinguishing token so the odd
# filenames (double spaces, "(1)", VOF vs F2F) still resolve.
# ---------------------------------------------------------------------------
CLASS_CONFIG = [
    # (match tokens (all must appear, case-insensitive), [class defs by table order])
    (["bsb50520"], [
        {"class": "Diploma Library Services - PTE Evening",     "qual": "BSB50520", "delivery": "Evening"},
        {"class": "Diploma Library Services - Face to Face",    "qual": "BSB50520", "delivery": "F2F"},
        {"class": "Diploma Library Services - Fulltime VOFF",   "qual": "BSB50520", "delivery": "VOFF"},
    ]),
    (["bsb40720"],            [{"class": "Cert IV VOCF",                "qual": "BSB40720", "delivery": ""}]),
    (["ict40120"],            [{"class": "Cert IV Programming",         "qual": "ICT40120", "delivery": ""}]),
    (["ict30120", "vof"],     [{"class": "Cert III General (VOFF)",     "qual": "ICT30120", "delivery": "VOFF"}]),
    (["ict30120", "f2f"],     [{"class": "Cert III General (F2F)",      "qual": "ICT30120", "delivery": "F2F"}]),
]

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


def auto_class_defs(fname, tables):
    """Derive class definitions from the filename + per-table headings when no
    explicit CLASS_CONFIG entry matches, so any new document 'just works'."""
    stem = os.path.splitext(fname)[0]
    qm = QUAL_RE.search(stem.upper())
    qual = qm.group(1) if qm else ""
    low = fname.lower()
    if "f2f" in low or "face to face" in low:
        mode = "F2F"
    elif "voff" in low or re.search(r"\bvof\b", low) or "online" in low:
        mode = "VOFF"
    elif "evening" in low or "evng" in low:
        mode = "Evening"
    else:
        mode = ""
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

    if class_defs is None:
        class_defs = auto_class_defs(fname, tables)
        issues.append({"file": fname, "level": "INFO",
                       "msg": f"Auto-detected {len(class_defs)} class(es): "
                              + "; ".join(d["class"] for d in class_defs)})
    elif len(tables) != len(class_defs):
        issues.append({"file": fname, "level": "WARN",
                       "msg": f"Expected {len(class_defs)} table(s) for configured classes "
                              f"but found {len(tables)}. Check class config / document structure."})

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
            # Semester precedence: stated in this row > carried from the row above
            # > inferred for the whole document > explicit caller fallback.
            # NEVER a hardcoded literal - see infer_doc_semester().
            if sem:
                sem_out, sem_src = sem, "row"
            elif last_sem:
                sem_out, sem_src = last_sem, "carried"
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
          "day", "channel_or_room",
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
        low = fname.lower()
        cfg = next((defs for toks, defs in CLASS_CONFIG if all(t in low for t in toks)), None)
        recs, iss = normalise_file(path, cfg, args.assume_semester)
        all_records.extend(recs); all_issues.extend(iss)
        tag = "" if cfg else "  (auto-detected classes)"
        sems = sorted({r["semester"] or "UNKNOWN" for r in recs})
        print(f"  {fname}: {len(recs)} session rows, {len(iss)} issue(s){tag}")
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
