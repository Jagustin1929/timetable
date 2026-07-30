---
inclusion: always
---

# Timetable Tool — Normalisation Conventions

Project-specific rules for the teacher-timetable pipeline (Agent 1 Normaliser and later phases).

## Delivery mode

- **"Online" MUST be normalised to `VOFF`.** Any occurrence of "Online" (in the delivery-mode
  field, the Teacher column, or a Unit-of-Competency note) is written out as `VOFF`.
- **"VOF" MUST be normalised to `VOFF`** (VOF and VOFF are the same thing).
- Recognised delivery modes after normalisation: `F2F`, `VOFF`, `Evening` (extend as needed).

## Teacher name matching (consolidation phase, not Agent 1)

- Teacher names that are **mostly the same but not exactly** are treated as the **same person**
  (e.g., "Anu Joshi" and "Anu" = same person; "Rima Andrews" and "R. Andrews" = same person).
- Canonicalisation happens in the **consolidation/extraction phase**, NOT in Agent 1 (Agent 1
  keeps names verbatim).
- **Every merge decision is recorded in the audit log**, and where a short/partial name could
  match **more than one** full name, the tool pauses/flags for confirmation rather than guessing.

## Source files → classes

| Source file                             | Class (own sheet)                              | Delivery |
|-----------------------------------------|------------------------------------------------|----------|
| BSB50520 Diploma Library Services combined | Diploma Library Services — PTE Evening      | Evening  |
| BSB50520 Diploma Library Services combined | Diploma Library Services — Face to Face     | F2F      |
| BSB50520 Diploma Library Services combined | Diploma Library Services — Fulltime VOFF    | VOFF     |
| BSB40720 Cert IV VOCF FTS2 2026         | Cert IV VOCF                                   | —        |
| ICT40120 Cert IV ... (combined)         | Cert IV Programming / Cert IV AI / Cert IV Data | —       |
| ICT30120 Cert III General VOF OUR       | Cert III General (online)                      | VOFF     |
| ICT30120 Cert III General F2F OUR       | Cert III General (face-to-face)                | F2F      |

- The **combined Diploma** is split into 3 classes by heading, in order:
  PTE Evening → Diploma Face to Face → Diploma Fulltime VOFF.
- The two **ICT30120** files are separate classes (not duplicates); a teacher may legitimately
  appear in both.

## Combined multi-course documents

- **The document decides how many classes it holds — never `CLASS_CONFIG`.** A combined
  file gains and loses course streams over time; a config entry is a naming hint only.
  Burying extra tables under one class name silently loses whole courses.
- One qualification may be delivered as several parallel **course streams**
  (ICT40120 → Programming, AI, Data). These arrive either as one file per stream or as
  one combined file with **one table per stream, under a heading naming the stream**.
- A stream **MUST produce the same class name whether it arrived standalone or combined**
  (`Cert IV Programming` either way). Class names are therefore built as
  `<qualification level> <stream>` — the level from the filename (`Cert IV`) or the AQF
  digit in the national code (`ICT40120` → 4 → `Cert IV`). If the names differ, uploads
  cannot supersede by class identity and every session is counted twice.
- Stream identification order: heading above the table → title row inside the table →
  unit-code prefixes taught (`ICTPRG`→Programming, `ICTAII`→AI, `ICTDBS`/`ICTDAT`→Data).
  Core units shared by all streams (`ICTICT`, `BSBXCS`, `BSBCRT`) must never decide it,
  and the winner needs a clear majority.
- **Adding a stream = one line in `STREAM_ALIASES`** (optionally its unit prefix in
  `STREAM_UNIT_PREFIXES`). No other change should be needed.
- Never silently merge two tables into one class: if two tables resolve to the same
  name, keep them separate and `WARN`. If a table's stream cannot be identified, name it
  by position and `WARN` — do not fold it into a neighbouring class.
- Match text for stream names on **word boundaries**, so `Date of Study` never reads as
  the Data stream and `Website` never as Web.

## Semester determination

- **The actual dates in the "Date of Study" column determine the semester and year.**
  Always derive it from those dates first; treat them as authoritative over any stated
  semester (log a `WARN` when they disagree). Boundary: months 1-6 = Semester 1,
  months 7-12 = Semester 2.
- **NEVER hardcode a semester default** (e.g. `"S2 2026"`) anywhere in the pipeline.
  Documents that state no semester must be marked `UNKNOWN`, not stamped with a guess.
- Precedence: **Date of Study dates** → row text → carried forward from row above →
  carried from preceding dated rows → inferred from filename/headings → explicit
  `--assume-semester` → `UNKNOWN`.
- Never read a truncated year (`03/12/202`) as a two-digit year; reject it and warn.
  Two dates may run together with no separator (`27/10/20271/12/2027`).
- When filtering to a target semester, rows with an **UNKNOWN** semester are **included**
  (they are undated, not "another semester") and logged as `SEMESTER-ASSUMED`. Only rows
  naming a *different* semester are excluded, logged as `SEMESTER-EXCLUDED`.
- If an UNKNOWN row's document pins a year (`... OUR 27` → `year_hint=2027`), exclude it
  from builds of a different year, logged as `SEMESTER-WRONG-YEAR`.
- Qualification codes (`ICT40120`, `BSB50520`) must be stripped before scanning text for
  years/semester numbers so their digits are never misread.

## Workload calculation

- Load policy lives in `timetable_tool/workload.py` — fractions, the 360-hour full load,
  the break rule and the status bands. Never hardcode these elsewhere.
- A full load is **360 delivery hours per semester**; expected hours = 360 x fraction.
- `Total hours = net session hours x number of weeks` for recurring delivery with
  numeric week ranges.
- **Unpaid break:** a session of MORE than 4 hours loses 30 minutes (9:00-2:30 = 5 hours,
  not 5.5). The threshold is 4 hours, not 3 — evening classes run 5:30-9:30pm with no
  break and must count in full. Apply it to the hours used for every contact-hour and
  workload figure, and show rostered / deduction / net separately so the number can be
  audited.
- Status bands: **ON TRACK** 90-110%, **UNDER** <90%, **OVER** >110% of expected hours.
- A teacher with no fraction defined must be reported (`WORKLOAD-NO-FRACTION`), never
  silently omitted from the workload check.

## Uploads and revisions

- Uploads live in `webapp_data/source_docs/`; never write to the repository's
  `source_docs/`, which is the reference copy.
- Support revising a **single** timetable mid-semester without re-uploading the whole set.
- Supersede a previous version **by class identity, not filename** — a revision often
  arrives renamed (`... combined.docx` -> `... combined 2.docx`), and keeping both would
  double-count every session and invent clashes.
- Always report what an upload added / updated / superseded / discarded / kept. Never
  replace or discard a timetable silently.

## Table columns

- Locate columns by **header text** (`Day`/`Day and Room`/`Day and Channel`,
  `Teacher`/`Teachers`/`Trainer`), falling back to fixed positions only for columns that
  cannot be identified — and raise a `WARN` when that happens.
- Never truncate rows to a fixed column count; a document with an extra column
  previously dropped the Teacher column silently.

## Agent 1 scope (structural normalisation only)

- No co-teaching split, no teacher renaming, no aggregation, no FTE. Co-teachers stay in one row.
- Sessions with **no teacher** trigger a pause/confirm (teacher may legitimately be "none").
- Session types: Teaching and Tutorial Support only.
- No Room data in source columns.
