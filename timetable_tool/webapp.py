#!/usr/bin/env python3
"""
Teacher Timetable Builder - local web app
=========================================
A no-dependency, no-internet local GUI:

    python3 timetable_tool/webapp.py        # opens http://127.0.0.1:8000

Flow:
  1. Upload the class-timetable Word files.
  2. The app normalises them and shows the semesters it found -> pick one.
  3. It builds the teacher timetables; download Excel / PDF / CSV / ZIP,
     or view them inline.

Identical results to the command-line pipeline (agent1 -> agent2 -> build_prototype
-> make_pdf).
"""
import os, sys, re, csv, subprocess, webbrowser, urllib.parse, socket
from collections import Counter
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

TOOLS = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(TOOLS)
WORK = os.path.join(REPO, "webapp_data")
SRC = os.path.join(WORK, "source_docs")
OUT = os.path.join(WORK, "output")
RESULTS_HTML = os.path.join(WORK, "results.html")
MASTER = os.path.join(OUT, "normalised_master.csv")
PY = sys.executable or "python3"

MIME = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".csv": "text/csv", ".pdf": "application/pdf", ".html": "text/html; charset=utf-8"}


# --------------------------- pipeline steps ---------------------------
def _run(label, cmd, log):
    log.append(f"$ {label}")
    try:
        p = subprocess.run(cmd, cwd=REPO, capture_output=True, text=True, timeout=300)
        if p.stdout.strip():
            log.append(p.stdout.strip())
        if p.stderr.strip():
            log.append("[stderr] " + p.stderr.strip())
        return p.returncode == 0
    except Exception as e:
        log.append(f"[error] {e}")
        return False


def run_normalise():
    os.makedirs(OUT, exist_ok=True)
    log = []
    ok = _run("Normalise Word documents",
              [PY, os.path.join(TOOLS, "agent1_normalise.py"), SRC, "--out", OUT], log)
    return ok and os.path.exists(MASTER), "\n".join(log)


def classes_in_docx(path):
    """The set of class names a single .docx resolves to.

    Used to supersede a previous version of a timetable even when the revision
    arrives under a different filename ("... combined.docx" ->
    "... combined 2.docx"). Matching on filename alone would leave the old file
    in place and double-count every session in it.

    A combined multi-course file returns one class per course stream, so uploading
    "ICT40120 Cert IV ... combined.docx" supersedes the older per-stream file
    ("... Programming OUR.docx") on the strength of the shared class name.
    """
    sys.path.insert(0, TOOLS)
    import agent1_normalise as a1
    try:
        recs, _issues = a1.normalise_file(path, a1.config_for(path))
    except Exception:
        return set()
    return {r["class"] for r in recs if r.get("class")}


def held_files():
    """Timetables currently held, as [(filename, sorted_classes), ...]."""
    os.makedirs(SRC, exist_ok=True)
    out = []
    for f in sorted(os.listdir(SRC)):
        if f.lower().endswith(".docx") and not f.startswith("~$"):
            out.append((f, sorted(classes_in_docx(os.path.join(SRC, f)))))
    return out


def apply_upload(docx, mode):
    """Write uploaded files into the working folder.

    mode "replace": start a clean set - discard everything held first.
    mode "update":  keep the timetables already held, but supersede any whose
                    classes the newly uploaded files also cover.

    Returns a list of human-readable report lines.
    """
    import shutil, tempfile
    os.makedirs(SRC, exist_ok=True)
    report = []

    if mode != "update":
        existing = [f for f, _c in held_files()]
        shutil.rmtree(SRC, ignore_errors=True)
        os.makedirs(SRC, exist_ok=True)
        for name, data in docx:
            with open(os.path.join(SRC, name), "wb") as fh:
                fh.write(data)
        for name, _d in docx:
            report.append(("added", name, ""))
        for f in existing:
            if f not in [n for n, _d in docx]:
                report.append(("discarded", f, "replaced set"))
        return report

    # --- update mode -------------------------------------------------------
    tmp = tempfile.mkdtemp()
    try:
        incoming = []
        for name, data in docx:
            p = os.path.join(tmp, name)
            with open(p, "wb") as fh:
                fh.write(data)
            incoming.append((name, p, classes_in_docx(p)))

        current = held_files()
        new_names = {n for n, _p, _c in incoming}
        new_classes = set().union(*[c for _n, _p, c in incoming]) if incoming else set()

        for fname, classes in current:
            if fname in new_names:
                continue                      # same filename -> overwritten below
            if classes and new_classes and set(classes) & new_classes:
                os.remove(os.path.join(SRC, fname))
                report.append(("superseded", fname,
                               "same class(es) as an uploaded file: "
                               + ", ".join(sorted(set(classes) & new_classes))))
            else:
                report.append(("kept", fname, ", ".join(classes)))

        for name, p, _c in incoming:
            verb = "updated" if name in {f for f, _c in current} else "added"
            shutil.copyfile(p, os.path.join(SRC, name))
            report.append((verb, name, ""))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return report


def _sem_key(label):
    m = re.search(r"S\s*([12])\s*(\d{4})", label)
    return (int(m.group(2)), int(m.group(1))) if m else (9999, 9)


def analyze_data():
    rows = list(csv.DictReader(open(MASTER, encoding="utf-8")))
    sem_counts = Counter(r["semester"].strip() for r in rows if r["semester"].strip())
    semesters = sorted(sem_counts.items(), key=lambda kv: _sem_key(kv[0]))
    default = sem_counts.most_common(1)[0][0] if sem_counts else ""
    classes = sorted(Counter(r["class"] for r in rows).items())
    # Rows whose document never stated a semester. These are included in
    # whichever semester is built, so the user needs to see them.
    undated = Counter(r["source_file"] for r in rows if not r["semester"].strip())
    return semesters, default, classes, undated


def run_extract(semester):
    sem_file = semester.replace(" ", "_")
    tt_xlsx = os.path.join(OUT, f"teacher_timetables_{sem_file}.xlsx")
    tt_pdf = os.path.join(OUT, f"teacher_timetables_{sem_file}.pdf")
    log = []
    ok = _run("Consolidate & extract teacher timetables",
              [PY, os.path.join(TOOLS, "agent2_extract.py"), MASTER,
               "--out", OUT, "--semester", semester], log)
    if ok:
        _run("Build results view",
             [PY, os.path.join(TOOLS, "build_prototype.py"),
              "--xlsx", tt_xlsx, "--dest", RESULTS_HTML, "--semester", semester,
              "--master", MASTER], log)
        _run("Render PDF timetables",
             [PY, os.path.join(TOOLS, "make_pdf.py"),
              "--xlsx", tt_xlsx, "--dest", tt_pdf, "--semester", semester], log)
    ok = os.path.exists(tt_xlsx)
    return ok, "\n".join(log), (f"teacher_timetables_{sem_file}.xlsx" if ok else None), \
        (f"teacher_timetables_{sem_file}.pdf" if os.path.exists(tt_pdf) else None)


# --------------------------- multipart parsing (stdlib) ---------------------------
def parse_multipart(body, boundary):
    fields, files = {}, []
    for part in body.split(b"--" + boundary.encode()):
        if part[:2] == b"\r\n":
            part = part[2:]
        if part[-2:] == b"\r\n":
            part = part[:-2]
        if not part or part == b"--" or b"\r\n\r\n" not in part:
            continue
        head, data = part.split(b"\r\n\r\n", 1)
        headers = head.decode("utf-8", "replace")
        fn = re.search(r'filename="([^"]*)"', headers)
        nm = re.search(r'name="([^"]*)"', headers)
        if fn and fn.group(1):
            files.append((os.path.basename(fn.group(1)), data))
        elif nm:
            fields[nm.group(1)] = data.decode("utf-8", "replace").strip()
    return fields, files


# --------------------------- HTML ---------------------------
CSS = """
:root{--ink:#1f2430;--mut:#6b7280;--line:#d7dbe3;--bg:#f3f5f9;--accent:#2f5bea;--ok:#0d7d5a;--warn:#c2410c}
*{box-sizing:border-box}body{margin:0;font:14.5px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:var(--ink);background:var(--bg)}
.wrap{max-width:880px;margin:0 auto;padding:26px 20px}
h1{font-size:20px;margin:0 0 4px}.sub{color:var(--mut);margin:0 0 20px}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:22px;margin-bottom:18px}
label{font-weight:600;display:block;margin:0 0 6px}
input[type=text],select{font:inherit;padding:9px 10px;border:1px solid var(--line);border-radius:8px;min-width:220px;background:#fff}
.drop{border:2px dashed var(--line);border-radius:10px;padding:26px;text-align:center;color:var(--mut);background:#fafbfe;margin:14px 0}
button{font:inherit;font-weight:600;background:var(--accent);color:#fff;border:0;border-radius:8px;padding:11px 20px;cursor:pointer}
a.dl.sec,button.sec{background:#fff;color:var(--accent);border:1px solid var(--accent)}
.hint{font-size:12.5px;color:var(--mut);margin-top:10px}
.dls{display:flex;gap:10px;flex-wrap:wrap;margin:6px 0 4px}
.dl{display:inline-block;border:1px solid var(--line);border-radius:8px;padding:9px 14px;text-decoration:none;color:var(--ink);background:#fff}
.dl b{color:var(--accent)}
ul.cls{margin:6px 0 0;padding-left:18px;color:var(--ink)}ul.cls li{margin:2px 0}
pre{background:#0f1729;color:#d7e0f5;padding:14px;border-radius:8px;overflow:auto;font-size:12.5px;max-height:280px}
iframe{width:100%;height:760px;border:1px solid var(--line);border-radius:10px;background:#fff}
.tag{display:inline-block;background:#eef2ff;color:var(--accent);border-radius:20px;padding:2px 10px;font-size:12px}
.err{border-color:#f0c0a8;background:#fff7ef;color:#7c3a12}
"""

def page(body):
    return ("<!DOCTYPE html><html><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<title>Teacher Timetable Builder</title><style>" + CSS +
            "</style></head><body><div class='wrap'>" + body + "</div></body></html>")


def upload_form():
    held = held_files()
    if held:
        rows = "".join(
            f"<li><b>{f}</b>"
            + (f" <span class='hint' style='margin:0'>&rarr; {', '.join(c)}</span>" if c else "")
            + "</li>" for f, c in held)
        current = f"""
      <div class="card">
        <label>Timetables currently held ({len(held)})</label>
        <ul class="cls">{rows}</ul>
        <div class="hint">Uploading a revised copy of any of these replaces it -
        including when the revision has a different filename, as long as it covers
        the same class.</div>
        <div style="margin-top:12px">
          <a class="dl sec" href="/clear" onclick="return confirm('Remove all held timetables?')">Remove all</a>
        </div>
      </div>"""
        mode_choice = """
        <label style="margin-top:18px">What should happen to the timetables already held?</label>
        <label style="font-weight:400;display:block;margin:6px 0">
          <input type="radio" name="mode" value="update" checked>
          <b>Add or update</b> &mdash; keep the others, replace only what I upload now
          <span class="hint" style="margin:0 0 0 22px;display:block">Use this for a revised
          timetable part-way through the semester.</span>
        </label>
        <label style="font-weight:400;display:block;margin:6px 0">
          <input type="radio" name="mode" value="replace">
          <b>Start a fresh set</b> &mdash; discard everything held and use only what I upload now
          <span class="hint" style="margin:0 0 0 22px;display:block">Use this at the start of a
          new semester.</span>
        </label>"""
        sub = "Upload a revised timetable, or start a fresh set for a new semester."
    else:
        current = ""
        mode_choice = '<input type="hidden" name="mode" value="replace">'
        sub = "Upload your class-timetable Word files to begin."

    return page(f"""
      <h1>Teacher Timetable Builder</h1>
      <p class="sub">{sub}</p>
      {current}
      <form class="card" method="POST" action="/analyze" enctype="multipart/form-data">
        <div class="drop">
          <div>Choose the class-timetable <b>.docx</b> files (you can select several at once)</div>
          <input type="file" name="files" accept=".docx" multiple required>
        </div>
        {mode_choice}
        <div style="margin-top:16px"><button type="submit">Continue &rarr;</button></div>
        <div class="hint">Runs entirely on your computer. Nothing is uploaded or installed.</div>
      </form>""")


def upload_report_html(report):
    """Summarise what the upload changed, so nothing is replaced silently."""
    if not report:
        return ""
    order = {"added": 0, "updated": 1, "superseded": 2, "discarded": 3, "kept": 4}
    colour = {"added": "var(--ok)", "updated": "var(--ok)",
              "superseded": "var(--warn)", "discarded": "var(--warn)", "kept": "var(--mut)"}
    def row_html(verb, name, detail):
        col = colour.get(verb, "var(--ink)")
        extra = f" <span class='hint' style='margin:0'>({detail})</span>" if detail else ""
        return (f"<li><span style='color:{col};font-weight:600'>{verb}</span> "
                f"{name}{extra}</li>")

    items = "".join(row_html(v, n, d) for v, n, d
                    in sorted(report, key=lambda r: (order.get(r[0], 9), r[1])))
    return f"""<div class="card"><label>Files used for this build</label>
      <ul class="cls">{items}</ul></div>"""


def choose_page(semesters, default, classes, undated=None, report=None):
    opts = "".join(
        f'<option value="{urllib.parse.quote(s)}"{" selected" if s == default else ""}>{s} ({n} sessions)</option>'
        for s, n in semesters)
    cls = "".join(f"<li>{c} <span class='hint'>({n})</span></li>" for c, n in classes)
    undated = undated or {}
    if undated:
        items = "".join(f"<li>{f} <span class='hint'>({n} sessions)</span></li>"
                        for f, n in sorted(undated.items()))
        note = (f"""<div class="card"><label>Documents that do not state a semester</label>
        <p class="sub">These {sum(undated.values())} session(s) will be included in whichever
        semester you choose above, because the document itself gives no semester.</p>
        <ul class="cls">{items}</ul></div>""")
    else:
        note = ""
    return page(f"""
      <h1>Choose semester <span class="tag">step 2 of 2</span></h1>
      <p class="sub">Detected {len(classes)} class(es) across your files. Pick the semester to build.</p>
      <form class="card" method="POST" action="/generate">
        <label>Semester</label>
        <select name="semester">{opts or '<option>(none detected)</option>'}</select>
        <div style="margin-top:16px"><button type="submit">Generate teacher timetables</button>
        <a class="dl sec" href="/" style="margin-left:8px">&#8617; Start over</a></div>
      </form>
      {note}
      <div class="card"><label>Classes detected</label><ul class="cls">{cls}</ul></div>
      {upload_report_html(report)}""")


def results_page(semester, xlsx_name, pdf_name, log):
    def dl(name, label):
        return (f'<a class="dl" href="/dl?f={urllib.parse.quote(name)}"><b>&#8595;</b> {label}</a>'
                if name and os.path.exists(os.path.join(OUT, name)) else "")
    downloads = "".join([
        dl(pdf_name or "", "Teacher timetables (PDF)"),
        dl(xlsx_name or "", "Teacher timetables (Excel)"),
        dl("workload_summary.csv", "Workload summary (CSV)"),
        dl("normalised.xlsx", "Normalised data (Excel)"),
        '<a class="dl" href="/zip"><b>&#8595;</b> All files (ZIP)</a>',
    ])
    highlights = "\n".join(l for l in log.splitlines()
                           if re.search(r"CLASHES|NO teacher|adjustments applied|assignments|Distinct|Combined|teacher page", l, re.I))
    return page(f"""
      <h1>Done &#10003; <span class="tag">{semester}</span></h1>
      <p class="sub">Download below, or view the timetables inline.</p>
      <div class="card"><label>Downloads</label><div class="dls">{downloads}</div>
        <a class="dl sec" href="/">&#8617; Start over with different files</a></div>
      <div class="card"><label>Run summary</label><pre>{highlights or 'See full log.'}</pre>
        <details><summary class="hint">Full log</summary><pre>{log}</pre></details></div>
      <div class="card"><label>Results view</label><iframe src="/view"></iframe></div>""")


# --------------------------- HTTP handler ---------------------------
class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="text/html; charset=utf-8", extra=None):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _clear_held(self):
        import shutil
        shutil.rmtree(SRC, ignore_errors=True)
        os.makedirs(SRC, exist_ok=True)
        for f in (MASTER, RESULTS_HTML):
            try:
                os.remove(f)
            except OSError:
                pass

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path == "/":
            self._send(200, upload_form())
        elif u.path == "/clear":
            self._clear_held()
            self._send(303, "", extra={"Location": "/"})
        elif u.path == "/view":
            if os.path.exists(RESULTS_HTML):
                self._send(200, open(RESULTS_HTML, "rb").read())
            else:
                self._send(404, page("<p>No results yet.</p>"))
        elif u.path == "/dl":
            name = os.path.basename((urllib.parse.parse_qs(u.query).get("f") or [""])[0])
            path = os.path.join(OUT, name)
            if name and os.path.exists(path):
                ext = os.path.splitext(name)[1].lower()
                self._send(200, open(path, "rb").read(), MIME.get(ext, "application/octet-stream"),
                           {"Content-Disposition": f'attachment; filename="{name}"'})
            else:
                self._send(404, page("<p>File not found.</p>"))
        elif u.path == "/zip":
            import io, zipfile
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for f in (sorted(os.listdir(OUT)) if os.path.isdir(OUT) else []):
                    fp = os.path.join(OUT, f)
                    if os.path.isfile(fp):
                        z.write(fp, f)
            self._send(200, buf.getvalue(), "application/zip",
                       {"Content-Disposition": 'attachment; filename="teacher_timetables_all.zip"'})
        elif u.path == "/favicon.ico":
            self._send(204, b"")
        else:
            self._send(404, page("<p>Not found.</p>"))

    def _read_form(self):
        ctype = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        m = re.search(r"boundary=([^;]+)", ctype)
        if "multipart/form-data" in ctype and m:
            return parse_multipart(body, m.group(1).strip().strip('"'))
        return dict(urllib.parse.parse_qsl(body.decode("utf-8", "replace"))), []

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/analyze":
            fields, files = self._read_form()
            docx = [(n, d) for n, d in files if n.lower().endswith(".docx")]
            if not docx:
                self._send(400, page("<div class='card err'><b>No .docx files received.</b> "
                                     "<a href='/'>Try again</a>.</div>")); return
            mode = (fields.get("mode") or "replace").strip()
            report = apply_upload(docx, mode)
            ok, log = run_normalise()
            if not ok:
                self._send(200, page(f"<div class='card err'><b>Could not read the documents.</b> "
                                     f"<a href='/'>Start over</a></div>"
                                     f"<div class='card'><pre>{log}</pre></div>")); return
            semesters, default, classes, undated = analyze_data()
            self._send(200, choose_page(semesters, default, classes, undated, report))
        elif path == "/generate":
            if not os.path.exists(MASTER):
                self._send(303, "", extra={"Location": "/"}); return
            fields, _ = self._read_form()
            semester = urllib.parse.unquote((fields.get("semester") or "S2 2026")).strip()
            ok, log, xlsx_name, pdf_name = run_extract(semester)
            if not ok:
                self._send(200, page(f"<div class='card err'><b>Generation did not complete.</b> "
                                     f"<a href='/'>Start over</a></div>"
                                     f"<div class='card'><pre>{log}</pre></div>")); return
            self._send(200, results_page(semester, xlsx_name, pdf_name, log))
        else:
            self._send(404, page("<p>Not found.</p>"))

    def log_message(self, *a):
        pass


def free_port(start=8000):
    for p in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


def main():
    os.makedirs(WORK, exist_ok=True)
    port = int(os.environ.get("PORT", 0)) or free_port(8000)
    url = f"http://127.0.0.1:{port}"
    print(f"\n  Teacher Timetable Builder is running at:  {url}")
    print("  (leave this window open; press Ctrl+C to stop)\n")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()


if __name__ == "__main__":
    main()
