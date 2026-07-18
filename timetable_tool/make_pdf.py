#!/usr/bin/env python3
"""Render individual teacher timetables to a single PDF (one teacher per page),
reading the workbook produced by agent2_extract.py. Standard library only.

    python3 timetable_tool/make_pdf.py --xlsx output/teacher_timetables_S2_2026.xlsx \
        --dest output/teacher_timetables_S2_2026.pdf --semester "S2 2026"
"""
import os, re, sys, zipfile, html, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_util import SimplePDF

# column label -> width (pts). Sum ~= 770 (usable width on A4 landscape, 36pt margins)
COLS = [("Day", 58), ("Time", 92), ("Weeks", 62), ("Hrs", 34),
        ("Class", 176), ("Mode", 60), ("Type", 78), ("Unit / activity", 210)]
# source column name for each display column
SRC = {"Day": "Day", "Time": "Time", "Weeks": "Weeks", "Hrs": "Hrs/session",
       "Class": "Class", "Mode": "Delivery", "Type": "Session type",
       "Unit / activity": "Units / activity"}
NON_TEACHER = {"Summary", "Reconciliation", "Clashes", "Unassigned sessions",
               "Adjustments & audit"}
M = 36
CODE_RE = re.compile(r"[A-Z]{3}[A-Z]{0,4}\d{3}")


def read_xlsx(path):
    z = zipfile.ZipFile(path)
    names = [html.unescape(n) for n in re.findall(r'name="([^"]+)"', z.read("xl/workbook.xml").decode())]
    sheets = {}
    for i, name in enumerate(names, start=1):
        xml = z.read(f"xl/worksheets/sheet{i}.xml").decode()
        rows = []
        for rm in re.findall(r"<row[^>]*>(.*?)</row>", xml, re.S):
            rows.append([html.unescape(c) for c in re.findall(r"<t[^>]*>(.*?)</t>", rm, re.S)])
        sheets[name] = rows
    return sheets


def trunc(s, width, size=8.0):
    maxc = max(3, int(width / (size * 0.52)))
    return s if len(s) <= maxc else s[:maxc - 2] + ".."


def draw_page_header(pdf, teacher, semester, subtitle, continued=False):
    pdf.text(M, 30, 15, f"Teacher Timetable  -  {teacher}" + ("  (cont.)" if continued else ""), bold=True)
    pdf.text(M, 46, 9.5, f"Semester {semester}    |    {subtitle}")
    y = 62
    total_w = sum(w for _, w in COLS)
    pdf.fill_rect(M, y, total_w, 15, gray=0.90)
    x = M
    for name, w in COLS:
        pdf.text(x + 3, y + 11, 8.5, name, bold=True)
        x += w
    pdf.line(M, y + 15, M + total_w, y + 15, 0.8)
    return y + 15


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--dest", required=True)
    ap.add_argument("--semester", default="")
    args = ap.parse_args()

    sh = read_xlsx(args.xlsx)
    teachers = [n for n in sh if n not in NON_TEACHER]

    # summary lookup for the per-teacher subtitle
    summ = {}
    if "Summary" in sh and sh["Summary"]:
        hdr, *rows = sh["Summary"]
        for r in rows:
            d = dict(zip(hdr, r))
            summ[d.get("Teacher", "")] = d

    pdf = SimplePDF()
    total_w = sum(w for _, w in COLS)
    row_h = 16
    bottom = pdf.h - 28

    for t in teachers:
        data = sh[t]
        hdr = data[0] if data else []
        body = data[1:]
        s = summ.get(t, {})
        subtitle = (f"{s.get('# sessions','?')} sessions | {s.get('# classes','?')} classes | "
                    f"{s.get('# distinct units','?')} units | "
                    f"{s.get('Semester contact hrs','?')} contact hrs")
        pdf.add_page()
        y = draw_page_header(pdf, t, args.semester, subtitle)
        for row in body:
            d = dict(zip(hdr, row))
            if y + row_h > bottom:                      # overflow -> new page
                pdf.add_page()
                y = draw_page_header(pdf, t, args.semester, subtitle, continued=True)
            x = M
            for name, w in COLS:
                val = d.get(SRC[name], "")
                if name == "Unit / activity" and d.get("Co-teacher(s)"):
                    val = f"{val}  (with {d['Co-teacher(s)']})"
                pdf.text(x + 3, y + 11, 8.0, trunc(str(val), w - 4))
                x += w
            pdf.line(M, y + row_h, M + total_w, y + row_h, 0.3, gray=0.75)
            y += row_h
        if not body:
            pdf.text(M + 3, y + 14, 9, "(no sessions this semester)")

    os.makedirs(os.path.dirname(args.dest) or ".", exist_ok=True)
    with open(args.dest, "wb") as fh:
        fh.write(pdf.output())
    print(f"Wrote {os.path.abspath(args.dest)}  ({len(teachers)} teacher page(s))")


if __name__ == "__main__":
    main()
