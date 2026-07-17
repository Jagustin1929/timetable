#!/usr/bin/env python3
"""
Agent 2/3 - Consolidate + Extract Individual Teacher Timetables
===============================================================
Reads the normalised master from Agent 1 and:
  * filters to a TARGET semester,
  * splits co-teachers into individual teachers (each gets the FULL session),
  * applies fuzzy teacher-name matching (near-identical names = same person),
  * applies MANUAL ADJUSTMENTS (adjustments.py) with per-teacher week ranges,
    logging every change for audit,
  * merges "combined" Cert III tutorials (one physical session listed twice),
  * detects clashes (same teacher, overlapping day/time/weeks),
  * computes an audit-ready workload summary (contact hours over the semester),
  * writes one Excel workbook + CSV copies.

Standard library only (Excel via shared xlsx_util fallback).

Usage:
    python3 agent2_extract.py output/normalised_master.csv --out output --semester "S2 2026"
"""
import sys, os, re, csv, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_util import write_workbook
from adjustments import ADJUSTMENTS, MERGE_COMBINED_CERT3, COMBINED_LABEL

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_ORDER = {d: i for i, d in enumerate(WEEKDAYS)}
MODE_TOKENS = [("face to face", "F2F"), ("on campus", "F2F"), ("oncampus", "F2F"),
               ("f2f", "F2F"), ("online", "VOFF"), ("voff", "VOFF"), ("vof", "VOFF"),
               ("campus", "F2F")]
UNIT_CODE_RE = re.compile(r"\b([A-Z]{3}[A-Z]{0,4}\d{3})\b")


# --------------------------- teacher parsing ---------------------------
def split_teachers(raw):
    if not raw or not raw.strip():
        return []
    out = []
    for part in re.split(r";|&|,|\band\b", raw):
        p = part.strip()
        if not p:
            continue
        mode = None
        low = p.lower()
        for tok, m in MODE_TOKENS:
            if tok in low:
                mode = m
                p = re.sub(re.escape(tok), "", p, flags=re.I)
        name = re.sub(r"\s+", " ", p).strip(" -")
        if name:
            out.append((name, mode))
    seen = {}
    for n, m in out:
        if n not in seen or (seen[n] is None and m):
            seen[n] = m
    return list(seen.items())


def is_same_person(a, b):
    a1, b1 = a.lower(), b.lower()
    if a1 == b1:
        return True
    ta, tb = a1.split(), b1.split()
    if set(ta) and set(tb) and (set(ta) <= set(tb) or set(tb) <= set(ta)):
        return True
    if ta and tb and ta[0] == tb[0] and (len(ta) == 1 or len(tb) == 1):
        return True
    return False


def canonicalize(names):
    uniq = sorted(set(names), key=lambda s: (-len(s), s))
    groups = []
    for name in uniq:
        for g in groups:
            if is_same_person(name, g[0]):
                g[1].add(name); break
        else:
            groups.append([name, {name}])
    mapping, merges = {}, []
    for canon, members in groups:
        for m in members:
            mapping[m] = canon
            if m != canon:
                merges.append((m, canon))
    canons = [g[0] for g in groups]
    ambiguous = [(n, [c for c in canons if c != n and is_same_person(n, c)])
                 for n in uniq]
    ambiguous = [(n, ms) for n, ms in ambiguous if len(ms) > 1]
    return mapping, merges, ambiguous


# --------------------------- time / weeks ---------------------------
def _parse_tok(tok):
    tok = tok.lower().replace("\u00a0", " ")
    ap = "pm" if "pm" in tok else ("am" if "am" in tok else None)
    m = re.search(r"(\d{1,2})(?::(\d{2}))?", tok)
    return (int(m.group(1)), int(m.group(2) or 0), ap) if m else None


def _to24(h, mm, ap):
    if ap == "pm" and h != 12:
        h += 12
    if ap == "am" and h == 12:
        h = 0
    return h * 60 + mm


def _infer(h, ap):
    return ap if ap else ("am" if 7 <= h <= 11 else "pm")


def time_span_minutes(start, end):
    """Return (start_min, end_min) in 24h, or None."""
    if not end:
        times = [t for t in re.findall(r"\d{1,2}(?::\d{2})?\s*(?:am|pm)?", start or "") if re.search(r"\d", t)]
        if len(times) >= 2:
            start, end = times[0], times[1]
    ps, pe = _parse_tok(start or ""), _parse_tok(end or "")
    if not ps or not pe:
        return None
    s = _to24(ps[0], ps[1], _infer(ps[0], ps[2]))
    e = _to24(pe[0], pe[1], _infer(pe[0], pe[2]))
    if e <= s:
        e += 720
    return s, e


def span_hours(start, end):
    span = time_span_minutes(start, end)
    if not span:
        return None
    dur = (span[1] - span[0]) / 60.0
    return round(dur, 2) if 0 < dur <= 12 else None


def parse_weeks(w):
    weeks = set()
    for a, b in re.findall(r"(\d+)\s*-\s*(\d+)", w):
        a, b = int(a), int(b)
        if 1 <= a <= 40 and a <= b <= 40:
            weeks.update(range(a, b + 1))
    if not re.search(r"\d+\s*-\s*\d+", w):
        for s in re.findall(r"\b(\d+)\b", w):
            if 1 <= int(s) <= 40:
                weeks.add(int(s))
    return weeks


def weeks_label(wkset):
    if not wkset:
        return ""
    ws = sorted(wkset); out = []; a = p = ws[0]
    for x in ws[1:]:
        if x == p + 1:
            p = x; continue
        out.append((a, p)); a = p = x
    out.append((a, p))
    return ", ".join(f"{a}-{b}" if a != b else f"{a}" for a, b in out)


def overlaps(a0, a1, b0, b1):
    return a0 < b1 and b0 < a1


# --------------------------- adjustments ---------------------------
def match_session(b, spec):
    if "class_contains" in spec and spec["class_contains"].lower() not in b["class"].lower():
        return False
    if "day" in spec and spec["day"] != b["day"]:
        return False
    if "time_start" in spec and not b["t_start"].startswith(spec["time_start"]):
        return False
    if "unit" in spec and spec["unit"].lower() not in b["units"].lower():
        return False
    return True


def apply_adjustments(bases, log):
    for adj in ADJUSTMENTS:
        matched = [b for b in bases if match_session(b, adj["match"])]
        if len(matched) != 1:
            log.append(["ADJ-WARN", f"{adj['id']}: matched {len(matched)} sessions "
                                     f"(expected 1) - {adj['match']}"])
        for b in matched:
            op = adj["op"]; who = adj.get("teacher")
            if op == "remove_teacher":
                b["pairs"] = [(n, m) for n, m in b["pairs"] if n != who]
                b["tweeks"].pop(who, None)
            elif op == "set_teacher_weeks":
                if who not in [n for n, _ in b["pairs"]]:
                    log.append(["ADJ-WARN", f"{adj['id']}: {who} not on matched session"])
                b["tweeks"][who] = parse_weeks(adj["weeks"])
            elif op == "add_teacher":
                if who not in [n for n, _ in b["pairs"]]:
                    b["pairs"].append((who, adj.get("mode")))
                b["tweeks"][who] = parse_weeks(adj["weeks"])
            log.append(["ADJ-APPLIED", f"{adj['id']} [{op}] {who} on "
                                       f"[{b['class']}] {b['day']} {b['time']} "
                                       f"({adj['match'].get('unit','')}): {adj['reason']}"])


# --------------------------- main ---------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("master_csv")
    ap.add_argument("--out", default="output")
    ap.add_argument("--semester", default="S2 2026")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    TARGET = args.semester
    audit = []

    rows = list(csv.DictReader(open(args.master_csv, encoding="utf-8")))
    in_sem = [r for r in rows if (r["semester"] or "").strip() == TARGET]

    # ---- base sessions (one per source row) ----
    bases = []
    for r in in_sem:
        wk = parse_weeks(r["weeks_raw"])
        bases.append({
            "class": r["class"], "day": r["day"], "day_order": DAY_ORDER.get(r["day"], 9),
            "t_start": r["time_start"], "t_end": r["time_end"],
            "time": f'{r["time_start"]}-{r["time_end"]}'.strip("-"),
            "hours": span_hours(r["time_start"], r["time_end"]),
            "wkset": wk, "weeks_raw": r["weeks_raw"],
            "session_type": r["session_type"], "units": r["units"] or r["notes"],
            "delivery": r["delivery"],
            "src": f'{r["source_file"]} t{r["src_table"]} r{r["src_row"]}',
            "pairs": split_teachers(r["teachers"]), "tweeks": {},
        })

    apply_adjustments(bases, audit)

    # ---- expand to per-teacher sessions (each with its own week set) ----
    all_names, unassigned, sessions = [], [], []
    for b in bases:
        if not b["pairs"]:
            unassigned.append(b); continue
        for name, mode in b["pairs"]:
            all_names.append(name)
            wkset = b["tweeks"].get(name, b["wkset"])
            sessions.append({**{k: b[k] for k in
                              ("class", "day", "day_order", "t_start", "t_end", "time",
                               "hours", "session_type", "units", "delivery", "src")},
                             "teacher_raw": name, "mode": mode or b["delivery"],
                             "wkset": wkset, "wkcount": len(wkset),
                             "coteachers": [n for n, _ in b["pairs"] if n != name]})

    mapping, merges, ambiguous = canonicalize(all_names)
    for s in sessions:
        s["teacher"] = mapping.get(s["teacher_raw"], s["teacher_raw"])

    # ---- merge combined Cert III tutorials (one physical session, listed twice) ----
    if MERGE_COMBINED_CERT3:
        by_key = {}
        drop = []
        for s in sessions:
            if "combined" in s["units"].lower() and "Cert III" in s["class"]:
                key = (s["teacher"], s["day"], s["t_start"], s["t_end"], s["session_type"])
                if key in by_key:
                    keep = by_key[key]
                    keep["class"] = COMBINED_LABEL
                    keep["wkset"] |= s["wkset"]; keep["wkcount"] = len(keep["wkset"])
                    drop.append(id(s))
                    audit.append(["COMBINED-MERGE",
                                  f"{s['teacher']} {s['day']} {s['time']} tutorial: merged "
                                  f"F2F+VOFF into one ({COMBINED_LABEL}), counted once."])
                else:
                    by_key[key] = s
        sessions = [s for s in sessions if id(s) not in drop]

    teachers = sorted({s["teacher"] for s in sessions})

    # ---- per-teacher sheets + clash detection ----
    T_HEADER = ["Day", "Time", "Weeks", "Hrs/session", "Contact hrs (x weeks)",
                "Class", "Delivery", "Session type", "Units / activity", "Co-teacher(s)"]
    per_sheet, per_hours, clashes = {}, {}, []
    for t in teachers:
        mine = sorted([s for s in sessions if s["teacher"] == t],
                      key=lambda s: (s["day_order"], s["t_start"]))
        rows_out, total = [], 0.0
        for s in mine:
            contact = round((s["hours"] or 0) * s["wkcount"], 2) if s["hours"] else ""
            if isinstance(contact, float):
                total += contact
            rows_out.append([s["day"], s["time"], weeks_label(s["wkset"]),
                             s["hours"] if s["hours"] is not None else "?", contact,
                             s["class"], s["mode"] or "", s["session_type"],
                             s["units"], ", ".join(s["coteachers"])])
        per_hours[t] = round(total, 2)
        per_sheet[t] = rows_out
        for i in range(len(mine)):
            for j in range(i + 1, len(mine)):
                a, b = mine[i], mine[j]
                if not a["day"] or a["day"] != b["day"] or a["class"] == b["class"]:
                    continue
                sa, sb = time_span_minutes(a["t_start"], a["t_end"]), time_span_minutes(b["t_start"], b["t_end"])
                if not sa or not sb:
                    continue
                wko = a["wkset"] & b["wkset"]
                if overlaps(sa[0], sa[1], sb[0], sb[1]) and wko:
                    clashes.append([t, a["day"], f'{a["time"]} vs {b["time"]}',
                                    weeks_label(wko), a["class"], b["class"]])

    # ---- summary ----
    summary_header = ["Teacher", "# sessions", "# classes", "# distinct units",
                      "Weekly contact hrs", "Semester contact hrs", "Classes"]
    summary_rows = []
    for t in teachers:
        mine = [s for s in sessions if s["teacher"] == t]
        classes = sorted({s["class"] for s in mine})
        units = {u for s in mine for u in UNIT_CODE_RE.findall(s["units"])}
        weekly = round(sum((s["hours"] or 0) for s in mine), 2)
        summary_rows.append([t, len(mine), len(classes), len(units),
                             weekly, per_hours[t], "; ".join(classes)])

    # ---- reconciliation + audit ----
    recon = [
        ["Target semester", TARGET],
        ["Source rows in semester", len(in_sem)],
        ["Teacher-session assignments (after adjustments)", len(sessions)],
        ["Distinct teachers", len(teachers)],
        ["Sessions with NO teacher (excluded)", len(unassigned)],
        ["Manual adjustments applied", sum(1 for a in audit if a[0] == "ADJ-APPLIED")],
        ["Combined tutorials merged", sum(1 for a in audit if a[0] == "COMBINED-MERGE")],
        ["Name-merge decisions", len(merges)],
        ["Ambiguous name matches", len(ambiguous)],
        ["Clashes detected", len(clashes)],
    ]
    for m, c in merges:
        audit.append(["NAME-MERGE", f"'{m}' -> '{c}'"])
    for name, ms in ambiguous:
        audit.append(["AMBIGUOUS-NAME", f"'{name}' could be: {', '.join(ms)}"])
    for u in unassigned:
        audit.append(["NO-TEACHER", f'[{u["class"]}] {u["day"]} {u["time"]} '
                                    f'{u["session_type"]}: {u["units"]}'])
    for s in sessions:
        if s["hours"] is None:
            audit.append(["HOURS?", f'[{s["teacher"]}] {s["class"]} {s["day"]} {s["time"]}'])

    # ---- workbook ----
    book = [("Summary", summary_header, summary_rows),
            ("Reconciliation", ["Metric", "Value"], recon)]
    if clashes:
        book.append(("Clashes", ["Teacher", "Day", "Times", "Overlap weeks", "Class A", "Class B"], clashes))
    book.append(("Unassigned sessions",
                 ["Class", "Day", "Time", "Weeks", "Session type", "Units / activity"],
                 [[u["class"], u["day"], u["time"], u["weeks_raw"], u["session_type"], u["units"]]
                  for u in unassigned]))
    book.append(("Adjustments & audit", ["type", "detail"], audit))
    for t in teachers:
        book.append((t, T_HEADER, per_sheet[t]))

    xlsx_path = os.path.join(args.out, f"teacher_timetables_{TARGET.replace(' ', '_')}.xlsx")
    engine = write_workbook(xlsx_path, book)
    with open(os.path.join(args.out, "workload_summary.csv"), "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh); w.writerow(summary_header); w.writerows(summary_rows)

    # ---- console ----
    print("=" * 74)
    print(f"TEACHER TIMETABLE EXTRACTION - semester {TARGET}")
    print("=" * 74)
    for k, v in recon:
        print(f"  {k:50} {v}")
    print("\nAdjustments applied:")
    for a in audit:
        if a[0] == "ADJ-APPLIED":
            print("   -", a[1])
    print("\nWorkload summary:")
    print(f"  {'Teacher':16}{'sess':>5}{'cls':>4}{'units':>6}{'wk hrs':>8}{'sem hrs':>9}")
    for r in summary_rows:
        print(f"  {r[0]:16}{r[1]:>5}{r[2]:>4}{r[3]:>6}{r[4]:>8}{r[5]:>9}")
    print(f"\nCLASHES remaining: {len(clashes)}")
    for c in clashes:
        print("   -", c)
    print("\nWrote:", xlsx_path, f"(via {engine})")


if __name__ == "__main__":
    main()
