#!/usr/bin/env python3
"""
Teacher Timetable Builder - local web app
=========================================
A no-dependency, no-internet local GUI. Run it, open the page in your browser,
upload the class-timetable Word files, click Generate, then download the Excel
and view the results.

    python3 timetable_tool/webapp.py         # then open http://127.0.0.1:8000

It orchestrates the same three scripts you'd run by hand (agent1 -> agent2 ->
build_prototype), so results are identical to the command-line pipeline.
"""
import os, sys, re, subprocess, webbrowser, urllib.parse, socket
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

TOOLS = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(TOOLS)
WORK = os.path.join(REPO, "webapp_data")
SRC = os.path.join(WORK, "source_docs")
OUT = os.path.join(WORK, "output")
RESULTS_HTML = os.path.join(WORK, "results.html")
PY = sys.executable or "python3"

MIME = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".csv": "text/csv", ".html": "text/html; charset=utf-8"}


# --------------------------- pipeline ---------------------------
def run_pipeline(semester):
    """Run agent1 -> agent2 -> build_prototype against WORK. Returns (ok, log, xlsx_name)."""
    os.makedirs(OUT, exist_ok=True)
    sem_file = semester.replace(" ", "_")
    tt_xlsx = os.path.join(OUT, f"teacher_timetables_{sem_file}.xlsx")
    steps = [
        ("Normalise Word documents",
         [PY, os.path.join(TOOLS, "agent1_normalise.py"), SRC, "--out", OUT]),
        ("Consolidate & extract teacher timetables",
         [PY, os.path.join(TOOLS, "agent2_extract.py"),
          os.path.join(OUT, "normalised_master.csv"), "--out", OUT, "--semester", semester]),
        ("Build results view",
         [PY, os.path.join(TOOLS, "build_prototype.py"),
          "--xlsx", tt_xlsx, "--dest", RESULTS_HTML, "--semester", semester]),
    ]
    log = []
    for label, cmd in steps:
        log.append(f"$ {label}")
        try:
            p = subprocess.run(cmd, cwd=REPO, capture_output=True, text=True, timeout=300)
            log.append(p.stdout.strip())
            if p.stderr.strip():
                log.append("[stderr] " + p.stderr.strip())
            if p.returncode != 0:
                return False, "\n".join(log), None
        except Exception as e:
            log.append(f"[error] {e}")
            return False, "\n".join(log), None
    ok = os.path.exists(tt_xlsx)
    return ok, "\n".join(log), (f"teacher_timetables_{sem_file}.xlsx" if ok else None)


# --------------------------- multipart parsing (stdlib) ---------------------------
def parse_multipart(body, boundary):
    fields, files = {}, []
    marker = b"--" + boundary.encode()
    for part in body.split(marker):
        if part[:2] == b"\r\n":
            part = part[2:]
        if part[-2:] == b"\r\n":
            part = part[:-2]
        if not part or part == b"--":
            continue
        if b"\r\n\r\n" not in part:
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
PAGE_CSS = """
:root{--ink:#1f2430;--mut:#6b7280;--line:#d7dbe3;--bg:#f3f5f9;--accent:#2f5bea;--ok:#0d7d5a;--warn:#c2410c}
*{box-sizing:border-box}body{margin:0;font:14.5px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:var(--ink);background:var(--bg)}
.wrap{max-width:860px;margin:0 auto;padding:26px 20px}
h1{font-size:20px;margin:0 0 4px}.sub{color:var(--mut);margin:0 0 20px}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:22px;margin-bottom:18px}
label{font-weight:600;display:block;margin:0 0 6px}
input[type=text]{font:inherit;padding:8px 10px;border:1px solid var(--line);border-radius:8px;width:200px}
.drop{border:2px dashed var(--line);border-radius:10px;padding:26px;text-align:center;color:var(--mut);background:#fafbfe;margin:14px 0}
.drop input{margin-top:8px}
button{font:inherit;font-weight:600;background:var(--accent);color:#fff;border:0;border-radius:8px;padding:11px 20px;cursor:pointer}
button.sec{background:#fff;color:var(--accent);border:1px solid var(--accent)}
.hint{font-size:12.5px;color:var(--mut);margin-top:10px}
.dls{display:flex;gap:10px;flex-wrap:wrap;margin:6px 0 4px}
.dl{display:inline-block;border:1px solid var(--line);border-radius:8px;padding:9px 14px;text-decoration:none;color:var(--ink);background:#fff}
.dl b{color:var(--accent)}
pre{background:#0f1729;color:#d7e0f5;padding:14px;border-radius:8px;overflow:auto;font-size:12.5px;max-height:280px}
iframe{width:100%;height:760px;border:1px solid var(--line);border-radius:10px;background:#fff}
.tag{display:inline-block;background:#eef2ff;color:var(--accent);border-radius:20px;padding:2px 10px;font-size:12px}
.err{border-color:#f0c0a8;background:#fff7ef;color:#7c3a12}
"""

def page(body):
    return ("<!DOCTYPE html><html><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<title>Teacher Timetable Builder</title><style>" + PAGE_CSS +
            "</style></head><body><div class='wrap'>" + body + "</div></body></html>")


def upload_form():
    return page(f"""
      <h1>Teacher Timetable Builder</h1>
      <p class="sub">Upload your class-timetable Word files, then generate individual teacher timetables.</p>
      <form class="card" method="POST" action="/run" enctype="multipart/form-data">
        <label>Semester</label>
        <input type="text" name="semester" value="S2 2026">
        <div class="drop">
          <div>Choose the class timetable <b>.docx</b> files (you can select several at once)</div>
          <input type="file" name="files" accept=".docx" multiple required>
        </div>
        <button type="submit">Generate teacher timetables</button>
        <div class="hint">Runs entirely on your computer. No files leave your machine; nothing is installed.</div>
      </form>
      <div class="card">
        <span class="tag">How it works</span>
        <p class="hint">Word files &rarr; normalised data &rarr; per-teacher timetables + workload summary +
        clash/audit report. The combined Diploma is split into its classes automatically, and the manual
        scheduling rules in <code>adjustments.py</code> are applied.</p>
      </div>""")


def results_page(semester, xlsx_name, log):
    def dl(name, label):
        p = os.path.join(OUT, name)
        if not os.path.exists(p):
            return ""
        return f'<a class="dl" href="/dl?f={urllib.parse.quote(name)}"><b>&#8595;</b> {label}</a>'
    downloads = "".join([
        dl(xlsx_name or "", "Teacher timetables (Excel)"),
        dl("workload_summary.csv", "Workload summary (CSV)"),
        dl("normalised.xlsx", "Normalised data (Excel)"),
        dl("normalised_master.csv", "Normalised master (CSV)"),
        '<a class="dl" href="/zip"><b>&#8595;</b> All files (ZIP)</a>',
    ])
    highlights = "\n".join(l for l in log.splitlines()
                           if re.search(r"CLASHES|NO teacher|adjustments applied|assignments|Distinct teachers|Combined", l, re.I))
    return page(f"""
      <h1>Done &#10003; <span class="tag">{semester}</span></h1>
      <p class="sub">Your teacher timetables are ready. Download below, or view them inline.</p>
      <div class="card">
        <label>Downloads</label>
        <div class="dls">{downloads or '<span class="hint">No output produced - see log below.</span>'}</div>
        <a class="dl sec" href="/">&#8617; Start over / upload different files</a>
      </div>
      <div class="card">
        <label>Run summary</label>
        <pre>{highlights or 'See full log below.'}</pre>
        <details><summary class="hint">Full log</summary><pre>{log}</pre></details>
      </div>
      <div class="card">
        <label>Results view</label>
        <iframe src="/view"></iframe>
      </div>""")


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

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path == "/":
            self._send(200, upload_form())
        elif u.path == "/view":
            if os.path.exists(RESULTS_HTML):
                self._send(200, open(RESULTS_HTML, "rb").read())
            else:
                self._send(404, page("<p>No results yet.</p>"))
        elif u.path == "/dl":
            q = urllib.parse.parse_qs(u.query)
            name = os.path.basename((q.get("f") or [""])[0])
            path = os.path.join(OUT, name)
            if name and os.path.exists(path):
                ext = os.path.splitext(name)[1].lower()
                self._send(200, open(path, "rb").read(),
                           MIME.get(ext, "application/octet-stream"),
                           {"Content-Disposition": f'attachment; filename="{name}"'})
            else:
                self._send(404, page("<p>File not found.</p>"))
        elif u.path == "/zip":
            import io, zipfile
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for f in sorted(os.listdir(OUT)) if os.path.isdir(OUT) else []:
                    fp = os.path.join(OUT, f)
                    if os.path.isfile(fp):
                        z.write(fp, f)
            self._send(200, buf.getvalue(), "application/zip",
                       {"Content-Disposition": 'attachment; filename="teacher_timetables_all.zip"'})
        elif u.path == "/favicon.ico":
            self._send(204, b"")
        else:
            self._send(404, page("<p>Not found.</p>"))

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != "/run":
            self._send(404, page("<p>Not found.</p>")); return
        ctype = self.headers.get("Content-Type", "")
        m = re.search(r"boundary=([^;]+)", ctype)
        if "multipart/form-data" not in ctype or not m:
            self._send(400, page("<p>Expected a file upload.</p>")); return
        boundary = m.group(1).strip().strip('"')
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        fields, files = parse_multipart(body, boundary)
        semester = (fields.get("semester") or "S2 2026").strip()
        docx = [(n, d) for n, d in files if n.lower().endswith(".docx")]
        if not docx:
            self._send(400, page("<div class='card err'><b>No .docx files received.</b> "
                                  "<a href='/'>Try again</a>.</div>")); return
        # fresh source dir
        import shutil
        shutil.rmtree(SRC, ignore_errors=True); os.makedirs(SRC, exist_ok=True)
        for name, data in docx:
            with open(os.path.join(SRC, name), "wb") as fh:
                fh.write(data)
        ok, log, xlsx_name = run_pipeline(semester)
        if not ok:
            self._send(200, page(f"<div class='card err'><b>Generation did not complete.</b> "
                                  f"<a href='/'>Start over</a></div>"
                                  f"<div class='card'><label>Log</label><pre>{log}</pre></div>"))
            return
        self._send(200, results_page(semester, xlsx_name, log))

    def log_message(self, *a):
        pass  # quiet


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
