# Teacher Timetable Builder

Generates individual, audit-ready **teacher timetables** and **workload summaries**
from class-timetable Word documents.

## Easiest way: the web app (no terminal commands)

```bash
python3 timetable_tool/webapp.py
```

This starts a small local app and opens it in your browser (or visit the
`http://127.0.0.1:8000` address it prints). Then just:

1. **Choose** your class-timetable `.docx` files (select several at once).
2. **Pick the semester** from the dropdown (auto-detected from your files) and click **Generate**.
3. **Download** the teacher timetables as **PDF** or **Excel**, the workload CSV, or
   **all files as a ZIP** — and view them inline.

Everything runs on your own machine — no files are uploaded anywhere and nothing
needs installing (pure Python standard library).

## Pipeline

1. **Agent 1 - Normaliser** (`timetable_tool/agent1_normalise.py`)
   Reads each `.docx` (stdlib only) and writes a normalised dataset
   (`output/normalised_master.csv` + `output/normalised.xlsx`).
   Classes are **auto-detected** (one Word table = one class; multi-class docs like
   the combined Diploma split automatically); `CLASS_CONFIG` optionally supplies
   curated names for known files. Any new document works without configuration.
   **Combined multi-course documents** (one table per course, each under a heading
   naming it - e.g. ICT40120 holds **Programming** and **AI/Data**) are split into
   one class per course. A course covering two subjects stays ONE course
   (`Cert IV AI/Data`). Add a subject with one line in `STREAM_ALIASES`.
   *Structural normalisation only - no co-teaching split, renaming, or FTE.*

2. **Agent 2/3 - Consolidate & Extract** (`timetable_tool/agent2_extract.py`)
   Filters to a target semester, splits co-teachers (each gets the full session),
   applies fuzzy teacher-name matching, applies **manual adjustments**
   (`timetable_tool/adjustments.py`), merges combined tutorials, detects clashes,
   and writes `output/teacher_timetables_<sem>.xlsx` + `output/workload_summary.csv`.

3. **Prototype wireframe** (`timetable_tool/build_prototype.py` -> `prototype/index.html`)
   A single-file clickable prototype populated with the real extracted data.

4. **PDF timetables** (`timetable_tool/make_pdf.py` + `pdf_util.py`)
   One page per teacher, written as a real PDF via a tiny pure-stdlib PDF writer
   (no external libraries) - ready to print or email.

The web app (`timetable_tool/webapp.py`) wraps all of the above behind a browser UI.

## Run from the command line (advanced / scripting)

```bash
python3 timetable_tool/agent1_normalise.py source_docs --out output
python3 timetable_tool/agent2_extract.py output/normalised_master.csv --out output --semester "S2 2026"
python3 timetable_tool/build_prototype.py
python3 timetable_tool/make_pdf.py --xlsx output/teacher_timetables_S2_2026.xlsx --dest output/teacher_timetables_S2_2026.pdf --semester "S2 2026"
```

No external dependencies required (Excel is written with a stdlib fallback;
`openpyxl` is used automatically if installed).

## Workload calculation (authoritative)

Defined in `timetable_tool/workload.py` — update it there each semester.

**A full teaching load is 360 delivery hours per semester.**

| Teacher | Fraction | Expected hrs |
|---|---|---|
| Graham Barber | 1.0 | 360 |
| Narelle Bell | 1.0 | 360 |
| Shaun Stummer | 0.9 | 324 |
| Jennie Agustin | 0.5 | 180 |
| Shanna Roper | 0.5 | 180 |
| Anu Joshi | 0.5 | 180 |
| Judi Stievenard | 0.3 | 108 |
| Martina Clark | 0.2 | 72 |
| Cathy Shay | 0.2 | 72 |
| Rima Andrews | 0.2 | 72 |

**Hours.** For recurring delivery with numeric week ranges:

```
Total hours = net session hours x number of weeks
```

**Unpaid break.** A session of **more than 4 hours** has **30 minutes** deducted, so
9:00–2:30 counts as **5** teaching hours, not 5.5. Sessions of 4 hours or less count in
full — evening classes run 5:30–9:30pm straight through with no break, so those 4 hours
are counted whole. The teacher sheets show the rostered hours, the deduction and the net
hours side by side so the figure can be audited.

**Status** — actual hours as a percentage of expected:

| Status | Band |
|---|---|
| **ON TRACK** | 90–110% |
| **UNDER** | below 90% |
| **OVER** | above 110% |

A teacher with no fraction defined gets no expected hours or status, and is reported in
the audit sheet under `WORKLOAD-NO-FRACTION` rather than being silently skipped.

## Uploading and revising timetables (web interface)

Uploaded files are kept in `webapp_data/source_docs/`. **The `source_docs/` folder in the
repository is never touched** — it stays as a reference copy.

The upload page lists the timetables currently held and offers two choices:

| Choice | Use when | Effect |
|---|---|---|
| **Add or update** | a revised timetable part-way through the semester | keeps everything already held; replaces only the classes you upload now |
| **Start a fresh set** | beginning of a new semester | discards everything held, uses only what you upload |

A revision supersedes the previous version **by class, not by filename**, so uploading
`... combined 2.docx` correctly replaces `... combined.docx` instead of leaving both in
place and double-counting every session. After each upload the app lists exactly what was
added, updated, superseded, discarded or kept, so nothing changes silently.

"Remove all" clears the held set.

## Semesters (important)

**The actual dates in the "Date of Study" column determine the semester and year.**
That column is the authoritative source — filenames and headings are only fallbacks
for rows that carry no usable date.

Semester boundary: **months 1–6 → Semester 1, months 7–12 → Semester 2.**
(Validated against every row in the source set that states its own semester: 47 agree,
0 disagree.)

Precedence:

1. **dates in the Date of Study column** — authoritative,
2. semester stated in the row itself (`Semester 1 2027`),
3. carried forward from the row above (merged cells),
4. carried from the preceding dated rows of the same table,
5. inferred from the filename/headings (e.g. `... FTS1 2027` → `S1 2027`),
6. an explicit `--assume-semester`,
7. otherwise **UNKNOWN**.

`semester_source` on every row records which of these was used, so the provenance is
always auditable. Where a row *states* a semester that disagrees with its dates, the
dates win and a `WARN` is raised. Rows whose dates span a semester boundary are filed
under the earliest date and flagged.

Date parsing handles the formats that occur in the real documents, including two dates
run together with no separator (`27/10/20271/12/2027`). A **truncated** year
(`03/12/202`) is rejected rather than read as year 2020, and raises a `WARN` naming the
row so it can be fixed at source.

Rows that end up UNKNOWN are **included in whichever semester you build** (logged as
`SEMESTER-ASSUMED`) — never silently dropped.

> Earlier versions stamped a hardcoded `S2 2026` on any document that stated no
> semester. Building any other semester silently dropped those documents entirely —
> and every teacher in them. Never reintroduce a hardcoded semester default.

## Conventions

See `.kiro/steering/timetable-normalisation.md`:
- `Online` and `VOF` normalise to `VOFF`.
- Near-identical teacher names are treated as the same person (consolidation phase).
- Manual scheduling overrides live in `timetable_tool/adjustments.py`, each with a
  reason, and every change is logged in the workbook's *Adjustments & audit* sheet.

## Layout

```
source_docs/     reference copy of the input Word documents - never written to
timetable_tool/  the pipeline scripts (incl. webapp.py - the GUI)
prototype/       clickable HTML wireframe
output/          results from the COMMAND LINE      (generated, not in git)
webapp_data/     uploads + results from the WEB APP (generated, not in git)
```

## Where to find your results

This catches people out, so it is worth being explicit:

| How you ran it | Uploads read from | Results written to |
|---|---|---|
| **Web app** (`webapp.py`) | `webapp_data/source_docs/` | `webapp_data/output/` |
| **Command line** | whatever folder you pass | `--out` folder (default `output/`) |

**Neither folder is committed to git**, deliberately. A checked-in `output/` looks
authoritative but goes stale the moment the source documents change, and reading it by
mistake means debugging data that has nothing to do with your upload.

The repository's `source_docs/` is a **reference copy only** - the web app never writes
to it, so its contents may be older than the files you have uploaded.

> **Opening `normalised_master.csv` in Excel:** the *Week of Study* column (`1-11`,
> `12-18`, `1-18`) is auto-converted to dates on open (`1-Nov`, `Dec-18`, `Jan-18`).
> That is Excel display only - but **do not save the file back**, or the mangled values
> are written in and every hours figure silently breaks, since weeks drive the totals.
> Open `normalised.xlsx` instead, or import the CSV with that column set to **Text**.
