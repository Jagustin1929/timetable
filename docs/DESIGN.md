# Design Document — Teacher Timetable Builder

**Status:** Implemented & validated on real data (S2 2026)
**Audience:** Reviewers, timetabling staff, future maintainers
**Last updated:** July 2026

---

## 1. Purpose

Convert **class** timetables (authored as Word documents) into accurate,
**per-teacher** timetables and **audit-ready workload summaries**.

The source documents describe *what each class does, when, and who teaches it*.
Staff instead need the inverse view — *what each individual teacher is doing
across every class* — plus defensible workload totals suitable for formal review.

## 2. Goals and non-goals

**Goals**
- Read the class timetable Word files with **no manual re-keying**.
- Produce one **individual timetable per teacher** for a chosen semester.
- Produce an **audit-ready workload summary** (contact hours per teacher).
- Surface every data-quality issue: **missing teachers, clashes, assumptions**.
- Make every **manual correction explicit, reasoned, and logged**.
- Run with **no external dependencies / no internet** (portable & reproducible).

**Non-goals (deliberately out of scope)**
- No automatic FTE / award-interpretation calculations.
- No live editing UI (the HTML artefact is a **prototype wireframe**, not the tool).
- No timetable *optimisation* — the tool reports clashes, it does not resolve them
  automatically (resolutions are human decisions, captured as adjustments).

## 3. Key requirements captured from stakeholders

| # | Requirement | Where handled |
|---|-------------|---------------|
| R1 | One class per Word table; the combined Diploma holds 3 classes | Agent 1 class config |
| R2 | Co-teaching → **both** teachers see the **full** session | Agent 2 expansion |
| R3 | Output is **per semester**; sessions may run <18 or >18 weeks | Semester filter + per-session week sets |
| R4 | Tutorial support shown **only if** a teacher is assigned | Unassigned sessions excluded from teacher sheets |
| R5 | A session with **no teacher** must pause/flag for confirmation | Audit "NO-TEACHER" + Unassigned sheet |
| R6 | `Online` and `VOF` normalise to `VOFF` | Delivery-mode normalisation |
| R7 | Near-identical teacher names = same person | Fuzzy canonicalisation (consolidation) |
| R8 | Everything correct, validated, and audit-ready at each step | Reconciliation + audit log |

Conventions are pinned in `.kiro/steering/timetable-normalisation.md`.

## 4. Architecture overview

A three-stage pipeline, each stage a standalone script with a clear contract.
Data flows through plain files (CSV/XLSX) so any stage can be inspected in isolation
— matching the stakeholder's mental model: *"put them into a single Excel document,
then extract from there."*

```
 5 Word docs (7 classes)
        │
        ▼
┌──────────────────────────┐
│ Agent 1 — Normaliser      │  agent1_normalise.py
│ Word → structured rows    │  (structural only: no split/rename/FTE)
└──────────────────────────┘
        │  normalised_master.csv  +  normalised.xlsx
        ▼
┌──────────────────────────┐
│ Agent 2/3 — Consolidate   │  agent2_extract.py
│ & Extract                 │  + adjustments.py  (manual overrides, logged)
│ • filter to semester      │
│ • split co-teachers       │
│ • fuzzy name matching     │
│ • apply adjustments       │
│ • merge combined sessions │
│ • detect clashes          │
│ • workload + reconcile    │
└──────────────────────────┘
        │  teacher_timetables_<sem>.xlsx  +  workload_summary.csv
        ▼
┌──────────────────────────┐
│ Prototype (wireframe)     │  build_prototype.py → prototype/index.html
│ clickable view of results │
└──────────────────────────┘
```

**Separation of concerns:** Agent 1 is *purely structural* — it never interprets
scheduling intent. All semantic decisions (who is the same person, who covers which
weeks, what counts once) live in Agent 2/3 and the adjustments file. This keeps the
fragile document-parsing logic isolated from the business logic.

## 5. Component design

### 5.1 Agent 1 — Normaliser (`agent1_normalise.py`)
- **`.docx` reading (stdlib only):** a `.docx` is a zip of XML. We parse
  `word/document.xml` with `xml.etree`, walking the body in document order so
  headings and tables keep their sequence.
- **Line-break fidelity:** text is extracted **preserving `<w:br>`/`<w:cr>` and
  paragraph boundaries**, so multiple units and multiple teachers packed into one
  cell split correctly (e.g. `Shaun StummerShanna Roper` → two names).
- **Class detection:** driven by a `CLASS_CONFIG` mapping filename → class(es).
  One Word **table = one class**; the combined Diploma yields 3 classes in table
  order (PTE Evening → Face to Face → Fulltime VOFF).
- **Merged-cell carry-forward:** blank Day/Time cells inherit the value above
  (Word vertical merges), reconstructing the intended grid.
- **Per-row normalisation:** day + semester/year, start/end time, unit codes+names,
  week ranges, dates, teachers (verbatim), delivery mode, and session type.
- **Output:** `normalised_master.csv` (one row per source session) and
  `normalised.xlsx` (Master + per-class + Audit sheets).

Agent 1 **keeps teachers verbatim** and does **not** split co-teaching — by design.

### 5.2 Agent 2/3 — Consolidate & Extract (`agent2_extract.py`)
Reads the normalised master and performs, in order:
1. **Semester filter** — keeps only rows whose parsed semester equals the target
   (default `S2 2026`). Necessary because the Diploma spans S2 2026 → S1 2028.
2. **Co-teacher split** — `Graham Barber; Graham Barber; Anu Joshi` → distinct
   teachers, de-duplicated, with any embedded delivery mode stripped and captured
   (`Shaun Stummer Online` → *Shaun Stummer* / `VOFF`).
3. **Fuzzy name matching** — near-identical names collapse to one canonical
   (longest form wins; subset/first-name rules). Ambiguous partial matches are
   **flagged, not guessed**.
4. **Adjustments** — manual overrides applied and logged (see §7).
5. **Combined-session merge** — one physical Cert III tutorial listed in both the
   F2F and VOFF documents is merged and **counted once**.
6. **Clash detection** — same teacher, same day, overlapping time **and** weeks,
   different classes.
7. **Workload + reconciliation** — see §8.

### 5.3 Adjustments (`adjustments.py`)
A data-only file of manual scheduling overrides. Each entry has a `match` (class /
day / time / unit), an `op`, and a plain-English `reason`. See §7.

### 5.4 Excel writer (`xlsx_util.py`)
`write_workbook(path, sheets)` uses `openpyxl` when present, otherwise a
**pure-stdlib** writer that emits a valid `.xlsx` (zip of XML with inline strings).
This is why the tool needs no installs and runs offline.

### 5.5 Prototype (`build_prototype.py` → `prototype/index.html`)
Reads the generated workbook and emits a **single self-contained HTML file** with
embedded JSON. Tabs: Import & pipeline, Teacher timetables (Mon–Fri grid), Workload,
Clashes & audit. It visualises the *real* output; it is a design artefact, not the
production interface.

## 6. Data model

**Normalised session row** (Agent 1 output — the contract between stages):

| Field | Meaning |
|-------|---------|
| `source_file`, `src_table`, `src_row` | provenance back to the exact Word cell |
| `class`, `qualification` | e.g. *Diploma … Fulltime VOFF*, `BSB50520` |
| `semester` | `S2 2026`, `S1 2027`, … (parsed; filtered later) |
| `day`, `channel_or_room` | weekday + channel/room token |
| `time_start`, `time_end` | tolerant of `-`, `–`, am/pm, no-space |
| `session_type` | Teaching / Tutorial Support / Self-directed learning |
| `units`, `unit_codes` | split unit code + name list |
| `notes` | non-unit activity text / delivery notes |
| `weeks_raw`, `dates_raw` | preserved verbatim for audit |
| `teachers` | **verbatim** (Agent 1 does not split) |
| `delivery` | normalised mode(s): `F2F` / `VOFF` / `Evening` |

**Teacher-session (Agent 2 internal):** the above expanded to one entry per teacher,
each carrying its **own week set** (`tweeks`) so a teacher can teach *part* of a
session's weeks (essential for the mid-semester hand-overs).

## 7. Adjustments framework (auditability)

Real timetables contain conflicts and shared sessions the source documents cannot
express. Rather than hand-editing data (which would be invisible to a reviewer),
every correction is a declarative record:

```python
{ "id": "ANU-2", "op": "set_teacher_weeks", "teacher": "Anu Joshi", "weeks": "1-12",
  "match": {"class_contains": "Cert IV Programming", "day": "Thursday",
            "time_start": "9:00", "unit": "ICTPRG302"},
  "reason": "Anu stops the Thursday ICT40120 session after wk12 (8 Nov 2026)." }
```

**Operations:** `remove_teacher`, `set_teacher_weeks`, `add_teacher`.
**Combined merge** is a separate rule for one-session-listed-twice tutorials.

Every applied adjustment (and every merge, name-match, missing teacher, and
uncomputable value) is written to the workbook's **Adjustments & audit** sheet.
A reviewer can therefore see exactly what changed, why, and what remains open.

## 8. Validation & audit strategy

- **Reconciliation:** `source rows = teacher-assigned + unassigned`; nothing is
  silently dropped. Reported every run.
- **Missing teachers:** never invented. Excluded from teacher sheets, listed in
  *Unassigned sessions*, and flagged for confirmation (R4/R5).
- **Clashes:** reported, not auto-resolved. Resolution is a human decision recorded
  as an adjustment.
- **Uncomputable hours / ambiguous names / unexpected table counts:** all raise
  explicit audit entries rather than failing silently.

**Contact-hours method:** session duration (with am/pm inference, e.g. `10:00–3:30`
→ 5.5h) × number of weeks in that teacher's week set for the session, summed per
teacher. Week counts come from the union of parsed ranges.

## 9. Output workbook structure (`teacher_timetables_<sem>.xlsx`)

`Summary` · `Reconciliation` · `Clashes` (if any) · `Unassigned sessions` ·
`Adjustments & audit` · then **one sheet per teacher** (Day, Time, Weeks,
Hrs/session, Contact hrs, Class, Delivery, Session type, Units, Co-teachers).

## 10. Key algorithms & the edge cases they solve

| Challenge (seen in real files) | Approach |
|--------------------------------|----------|
| Multiple units / teachers in one cell | preserve line breaks during extraction |
| Merged Day/Time cells | carry-forward from the row above |
| Header wording varies (*Day and Channel/Room/Day*) | match by **column position**, confirm by header |
| `Online`, `VOF` | normalise to `VOFF` |
| Duplicated names, `and`/`;` separators | split + de-duplicate |
| Same person, different spelling | fuzzy canonicalisation, ambiguity flagged |
| Multi-year Diploma | semester filter |
| One tutorial in two cohort docs | combined-merge, counted once |
| Teacher on only some weeks of a session | per-teacher week sets |
| am/pm-less times (`10:00-3:30`) | infer by hour, correct negative spans |

## 11. Assumptions & limitations

- Class↔file mapping is **auto-detected** (one Word table = one class, named from the
  preceding heading or the filename's qualification code). `CLASS_CONFIG` optionally
  overrides this to give curated names for known files, but new documents work with
  no configuration.
- Week numbers are the source of truth for coverage; **calendar dates are recorded
  as annotations** (the academic calendar has non-teaching breaks, so week↔date is
  not a fixed arithmetic).
- "on Campus" is treated as `F2F` (stated assumption).
- The prototype is illustrative; it is not wired to live editing.

## 12. Open items (tracked, not blocking)

- Remaining **unassigned** teaching gaps in the *Unassigned sessions* sheet await
  teacher confirmation (each becomes an adjustment).
- Optional future work: week-by-week grid view; GitHub Pages hosting of the
  prototype; direct `.docx` export of individual teacher timetables.

## 13. How to run

```bash
python3 timetable_tool/agent1_normalise.py source_docs --out output
python3 timetable_tool/agent2_extract.py output/normalised_master.csv --out output --semester "S2 2026"
python3 timetable_tool/build_prototype.py
```

No external dependencies required. `openpyxl` is used automatically if installed;
otherwise the stdlib writer produces the same workbook.
