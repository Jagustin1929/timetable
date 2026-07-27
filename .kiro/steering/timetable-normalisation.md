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

## Source files → classes (7 classes across 5 files)

| Source file                             | Class (own sheet)                              | Delivery |
|-----------------------------------------|------------------------------------------------|----------|
| BSB50520 Diploma Library Services combined | Diploma Library Services — PTE Evening      | Evening  |
| BSB50520 Diploma Library Services combined | Diploma Library Services — Face to Face     | F2F      |
| BSB50520 Diploma Library Services combined | Diploma Library Services — Fulltime VOFF    | VOFF     |
| BSB40720 Cert IV VOCF FTS2 2026         | Cert IV VOCF                                   | —        |
| ICT40120 Cert IV Programming OUR        | Cert IV Programming                            | —        |
| ICT30120 Cert III General VOF OUR       | Cert III General (online)                      | VOFF     |
| ICT30120 Cert III General F2F OUR       | Cert III General (face-to-face)                | F2F      |

- The **combined Diploma** is split into 3 classes by heading, in order:
  PTE Evening → Diploma Face to Face → Diploma Fulltime VOFF.
- The two **ICT30120** files are separate classes (not duplicates); a teacher may legitimately
  appear in both.

## Semester determination

- **NEVER hardcode a semester default** (e.g. `"S2 2026"`) anywhere in the pipeline.
  Documents that state no semester must be marked `UNKNOWN`, not stamped with a guess.
- Precedence: row text → carried forward from row above → inferred from
  filename/headings (`FTS1 2027` → `S1 2027`) → explicit `--assume-semester` → `UNKNOWN`.
- When filtering to a target semester, rows with an **UNKNOWN** semester are **included**
  (they are undated, not "another semester") and logged as `SEMESTER-ASSUMED`. Only rows
  naming a *different* semester are excluded, logged as `SEMESTER-EXCLUDED`.
- Qualification codes (`ICT40120`, `BSB50520`) must be stripped before scanning text for
  years/semester numbers so their digits are never misread.

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
