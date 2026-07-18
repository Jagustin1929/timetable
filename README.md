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
2. Set the **semester** (default `S2 2026`) and click **Generate**.
3. **Download** the teacher-timetables Excel / workload CSV, and view the results inline.

Everything runs on your own machine — no files are uploaded anywhere and nothing
needs installing (pure Python standard library).

## Pipeline

1. **Agent 1 - Normaliser** (`timetable_tool/agent1_normalise.py`)
   Reads each `.docx` (stdlib only), splits multi-class documents (the combined
   Diploma becomes 3 classes), and writes a normalised dataset
   (`output/normalised_master.csv` + `output/normalised.xlsx`).
   *Structural normalisation only - no co-teaching split, renaming, or FTE.*

2. **Agent 2/3 - Consolidate & Extract** (`timetable_tool/agent2_extract.py`)
   Filters to a target semester, splits co-teachers (each gets the full session),
   applies fuzzy teacher-name matching, applies **manual adjustments**
   (`timetable_tool/adjustments.py`), merges combined tutorials, detects clashes,
   and writes `output/teacher_timetables_<sem>.xlsx` + `output/workload_summary.csv`.

3. **Prototype wireframe** (`timetable_tool/build_prototype.py` -> `prototype/index.html`)
   A single-file clickable prototype populated with the real extracted data.

## Run from the command line (advanced / scripting)

```bash
python3 timetable_tool/agent1_normalise.py source_docs --out output
python3 timetable_tool/agent2_extract.py output/normalised_master.csv --out output --semester "S2 2026"
python3 timetable_tool/build_prototype.py
```

No external dependencies required (Excel is written with a stdlib fallback;
`openpyxl` is used automatically if installed).

## Conventions

See `.kiro/steering/timetable-normalisation.md`:
- `Online` and `VOF` normalise to `VOFF`.
- Near-identical teacher names are treated as the same person (consolidation phase).
- Manual scheduling overrides live in `timetable_tool/adjustments.py`, each with a
  reason, and every change is logged in the workbook's *Adjustments & audit* sheet.

## Layout

```
source_docs/     the 5 input Word documents (7 classes)
timetable_tool/  the pipeline scripts (incl. webapp.py - the GUI)
output/          generated CSV/XLSX
prototype/       clickable HTML wireframe
webapp_data/     scratch area used by the web app (ignored by git)
```
