#!/usr/bin/env python3
"""Build a single-file clickable HTML wireframe/prototype populated with the real
extracted S2 2026 data (reads the workbook produced by agent2_extract.py)."""
import os, re, json, zipfile, html

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "output")
XLSX = os.path.join(OUT, "teacher_timetables_S2_2026.xlsx")
DEST = os.path.join(HERE, "..", "prototype", "index.html")


def read_xlsx(path):
    z = zipfile.ZipFile(path)
    names = [html.unescape(n) for n in re.findall(r'name="([^"]+)"', z.read("xl/workbook.xml").decode())]
    sheets = {}
    for i, name in enumerate(names, start=1):
        xml = z.read(f"xl/worksheets/sheet{i}.xml").decode()
        rows = []
        for rm in re.findall(r"<row[^>]*>(.*?)</row>", xml, re.S):
            cells = re.findall(r"<t[^>]*>(.*?)</t>", rm, re.S)
            rows.append([html.unescape(c) for c in cells])
        sheets[name] = rows
    return sheets


def main():
    sh = read_xlsx(XLSX)
    non_teacher = {"Summary", "Reconciliation", "Clashes", "Unassigned sessions",
                   "Adjustments & audit"}
    teacher_names = [n for n in sh if n not in non_teacher]

    teachers = {}
    for t in teacher_names:
        hdr, *rows = sh[t]
        teachers[t] = [dict(zip(hdr, r)) for r in rows]

    def table(name):
        if name not in sh:
            return []
        hdr, *rows = sh[name]
        return [dict(zip(hdr, r)) for r in rows]

    files = [
        {"name": "BSB50520 Diploma Library Services combined.docx",
         "classes": ["Diploma Library Services - PTE Evening",
                     "Diploma Library Services - Face to Face",
                     "Diploma Library Services - Fulltime VOFF"]},
        {"name": "BSB40720 Cert IV VOCF FTS2 2026.docx", "classes": ["Cert IV VOCF"]},
        {"name": "ICT40120 Cert IV Programming OUR.docx", "classes": ["Cert IV Programming"]},
        {"name": "ICT30120 Cert III General VOF OUR.docx", "classes": ["Cert III General (VOFF)"]},
        {"name": "ICT30120 Cert III General F2F OUR.docx", "classes": ["Cert III General (F2F)"]},
    ]

    data = {
        "semester": "S2 2026",
        "files": files,
        "summary": table("Summary"),
        "reconciliation": {r["Metric"]: r["Value"] for r in table("Reconciliation")},
        "unassigned": table("Unassigned sessions"),
        "audit": table("Adjustments & audit"),
        "clashes": table("Clashes"),
        "teachers": teachers,
    }

    os.makedirs(os.path.dirname(DEST), exist_ok=True)
    with open(DEST, "w", encoding="utf-8") as fh:
        fh.write(HTML.replace("/*__DATA__*/", json.dumps(data)))
    print("Wrote", os.path.abspath(DEST))
    print("Teachers:", len(teachers), "| summary rows:", len(data["summary"]),
          "| unassigned:", len(data["unassigned"]), "| audit:", len(data["audit"]))


HTML = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Teacher Timetable Builder - Prototype</title>
<style>
:root{--ink:#1f2430;--mut:#6b7280;--line:#d7dbe3;--bg:#f3f5f9;--card:#fff;--accent:#2f5bea;
--f2f:#0d7d5a;--voff:#8a4fd0;--eve:#b7791f;--warn:#c2410c;--ok:#0d7d5a;}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:var(--ink);background:var(--bg)}
header{background:#fff;border-bottom:1px solid var(--line);padding:12px 20px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:5}
header h1{font-size:16px;margin:0}.wf{font-size:11px;color:var(--mut);border:1px dashed var(--line);border-radius:20px;padding:2px 10px}
.sem{margin-left:auto;font-size:12px;color:var(--mut)}.sem b{color:var(--ink)}
nav{display:flex;gap:2px;padding:0 20px;background:#fff;border-bottom:1px solid var(--line);position:sticky;top:49px;z-index:4;flex-wrap:wrap}
nav button{border:0;background:none;padding:12px 14px;font:inherit;color:var(--mut);cursor:pointer;border-bottom:2px solid transparent}
nav button.on{color:var(--accent);border-bottom-color:var(--accent);font-weight:600}
main{padding:20px;max-width:1180px;margin:0 auto}
.panel{display:none}.panel.on{display:block}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:16px}
h2{font-size:15px;margin:0 0 4px}.sub{color:var(--mut);font-size:12px;margin:0 0 14px}
.steps{display:flex;gap:10px;flex-wrap:wrap;margin:0 0 8px}.step{flex:1;min-width:150px;border:1px solid var(--line);border-radius:8px;padding:10px 12px;background:#fafbfe}
.step .n{width:20px;height:20px;border-radius:50%;background:var(--ok);color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:11px;margin-right:6px}
.drop{border:2px dashed var(--line);border-radius:10px;padding:26px;text-align:center;color:var(--mut);background:#fafbfe}
.filelist{margin-top:12px}.frow{display:flex;align-items:center;gap:10px;padding:9px 10px;border:1px solid var(--line);border-radius:8px;margin-bottom:8px;background:#fff}
.doc{width:26px;height:26px;border-radius:5px;background:#eef2ff;color:var(--accent);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-left:auto}
.chip{font-size:11px;border:1px solid var(--line);border-radius:20px;padding:2px 9px;color:var(--ink);background:#fff}
.grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}
.col h4{margin:0 0 8px;font-size:12px;color:var(--mut);text-align:center;text-transform:uppercase;letter-spacing:.04em}
.block{border:1px solid var(--line);border-left:4px solid var(--accent);border-radius:8px;padding:8px 9px;margin-bottom:8px;background:#fff}
.block .t{font-size:12px;color:var(--mut)}.block .u{font-size:12.5px;font-weight:600;margin:2px 0}
.block .m{display:flex;gap:6px;flex-wrap:wrap;margin-top:5px;align-items:center}
.tag{font-size:10.5px;border-radius:5px;padding:1px 7px;color:#fff}.tag.f2f{background:var(--f2f)}.tag.voff{background:var(--voff)}
.tag.eve{background:var(--eve)}.tag.tut{background:#475569}.tag.wk{background:#eef2ff;color:var(--accent);border:1px solid #dbe3ff}
.co{font-size:11px;color:var(--mut)}
.tlist{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px}
.tbtn{border:1px solid var(--line);background:#fff;border-radius:20px;padding:6px 13px;cursor:pointer;font:inherit}
.tbtn.on{background:var(--accent);color:#fff;border-color:var(--accent)}
.tbtn .h{font-size:11px;opacity:.8;margin-left:6px}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line)}
th{color:var(--mut);font-weight:600;font-size:11.5px;text-transform:uppercase;letter-spacing:.03em}
td.num,th.num{text-align:right}tr:hover td{background:#fafbfe}
.kpis{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px}
.kpi{flex:1;min-width:150px;border:1px solid var(--line);border-radius:10px;padding:14px 16px;background:#fff}
.kpi .v{font-size:24px;font-weight:700}.kpi .k{color:var(--mut);font-size:12px}
.kpi.good .v{color:var(--ok)}.kpi.warn .v{color:var(--warn)}
.badge{font-size:11px;border-radius:6px;padding:2px 8px;font-weight:600}
.badge.adj{background:#eef2ff;color:var(--accent)}.badge.merge{background:#e7f6ef;color:var(--ok)}
.badge.no{background:#fff2e8;color:var(--warn)}.badge.name{background:#f1f0fb;color:var(--voff)}
.auditrow{display:flex;gap:10px;padding:9px 0;border-bottom:1px solid var(--line);font-size:13px}
.hint{font-size:12px;color:var(--mut);margin-top:6px}
.callout{border:1px solid #ffe0c2;background:#fff7ef;border-radius:8px;padding:10px 12px;font-size:13px;color:#7c3a12;margin-bottom:12px}
</style></head><body>
<header><h1>Teacher Timetable Builder</h1><span class="wf">PROTOTYPE - wireframe</span>
<span class="sem">Semester <b id="semlbl"></b> - populated with real extracted data</span></header>
<nav id="nav"></nav>
<main>
 <section class="panel" id="p-upload">
  <div class="card"><h2>1 &middot; Import class timetables</h2>
   <p class="sub">Drop the Word timetable documents. Each file may contain more than one class.</p>
   <div class="drop">Drag &amp; drop .docx files here, or click to browse<div class="hint">(prototype - files below are the ones already processed)</div></div>
   <div class="filelist" id="files"></div>
  </div>
  <div class="card"><h2>2 &middot; Pipeline</h2>
   <div class="steps">
    <div class="step"><span class="n">1</span>Normalise (Word &rarr; structured)</div>
    <div class="step"><span class="n">2</span>Consolidate + name matching</div>
    <div class="step"><span class="n">3</span>Apply adjustments</div>
    <div class="step"><span class="n">4</span>Extract teacher timetables</div>
   </div>
   <div class="hint">All steps complete. Every class timetable was split by class; the combined Diploma became 3 classes.</div>
  </div>
 </section>

 <section class="panel" id="p-teachers">
  <div class="card">
   <h2>Individual teacher timetables</h2>
   <p class="sub">Select a teacher. Co-taught sessions appear in full for every teacher. Modes: <span class="tag f2f">F2F</span> <span class="tag voff">VOFF</span> <span class="tag eve">Evening</span></p>
   <div class="tlist" id="tlist"></div>
   <div id="thours" class="callout" style="display:none"></div>
   <div class="grid" id="tgrid"></div>
  </div>
 </section>

 <section class="panel" id="p-workload">
  <div class="card"><h2>Workload summary</h2><p class="sub">Audit-ready. Semester hours = session length &times; number of weeks.</p>
   <table id="wtab"></table></div>
 </section>

 <section class="panel" id="p-audit">
  <div class="kpis" id="kpis"></div>
  <div class="card"><h2>Clashes</h2><div id="clashes"></div></div>
  <div class="card"><h2>Sessions needing a teacher</h2>
   <p class="sub">The pipeline pauses on these - confirm a teacher or mark as intentionally unstaffed.</p>
   <table id="unassigned"></table></div>
  <div class="card"><h2>Adjustments &amp; audit log</h2>
   <p class="sub">Every manual override and merge, with its reason.</p>
   <div id="auditlog"></div></div>
 </section>
</main>
<script>
const DATA = /*__DATA__*/;
document.getElementById('semlbl').textContent = DATA.semester;
const tabs=[['upload','Import & pipeline'],['teachers','Teacher timetables'],['workload','Workload'],['audit','Clashes & audit']];
const nav=document.getElementById('nav');
tabs.forEach(([id,label],i)=>{const b=document.createElement('button');b.textContent=label;b.onclick=()=>show(id);b.id='nav-'+id;if(i===0)b.classList.add('on');nav.appendChild(b);});
function show(id){document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));document.getElementById('p-'+id).classList.add('on');
 document.querySelectorAll('nav button').forEach(b=>b.classList.remove('on'));document.getElementById('nav-'+id).classList.add('on');}
show('upload');

// files
const fl=document.getElementById('files');
DATA.files.forEach(f=>{const d=document.createElement('div');d.className='frow';
 d.innerHTML=`<span class="doc">W</span><div><div>${f.name}</div><div class="hint">${f.classes.length} class(es)</div></div>
 <div class="chips">${f.classes.map(c=>`<span class="chip">${c}</span>`).join('')}</div>`;fl.appendChild(d);});

// modes -> tag
function modeTags(mode,type){let h='';(mode||'').split('/').filter(Boolean).forEach(m=>{const c=m.toLowerCase().includes('voff')?'voff':m.toLowerCase().includes('f2f')?'f2f':'eve';h+=`<span class="tag ${c}">${m}</span>`;});
 if((type||'').toLowerCase().includes('tutorial'))h+=`<span class="tag tut">Tutorial</span>`;return h;}
const DAYS=['Monday','Tuesday','Wednesday','Thursday','Friday'];

// teacher list
const names=Object.keys(DATA.teachers).sort();
const tl=document.getElementById('tlist');
const hoursByT={};DATA.summary.forEach(r=>hoursByT[r.Teacher]=r['Semester contact hrs']);
names.forEach((n,i)=>{const b=document.createElement('button');b.className='tbtn'+(i===0?' on':'');
 b.innerHTML=`${n}<span class="h">${hoursByT[n]||''}h</span>`;b.onclick=()=>pickTeacher(n,b);tl.appendChild(b);});
function pickTeacher(n,btn){document.querySelectorAll('.tbtn').forEach(x=>x.classList.remove('on'));btn.classList.add('on');
 const rows=DATA.teachers[n];const sum=DATA.summary.find(r=>r.Teacher===n)||{};
 const co=document.getElementById('thours');co.style.display='block';
 co.innerHTML=`<b>${n}</b> &middot; ${rows.length} sessions across ${sum['# classes']||'?'} classes &middot; ${sum['# distinct units']||'?'} units &middot; <b>${sum['Semester contact hrs']||'?'}h</b> semester contact`;
 const grid=document.getElementById('tgrid');grid.innerHTML='';
 DAYS.forEach(day=>{const col=document.createElement('div');col.className='col';col.innerHTML=`<h4>${day}</h4>`;
  rows.filter(r=>r.Day===day).forEach(r=>{const bl=document.createElement('div');bl.className='block';
   bl.innerHTML=`<div class="t">${r.Time}</div><div class="u">${r['Units / activity']||r['Session type']}</div>
   <div class="t">${r.Class}</div>
   <div class="m"><span class="tag wk">Wk ${r.Weeks}</span>${modeTags(r.Delivery,r['Session type'])}<span class="co">${r['Contact hrs (x weeks)']?('&middot; '+r['Contact hrs (x weeks)']+'h'):''}</span></div>
   ${r['Co-teacher(s)']?`<div class="co">with ${r['Co-teacher(s)']}</div>`:''}`;col.appendChild(bl);});
  grid.appendChild(col);});}
pickTeacher(names[0],document.querySelector('.tbtn'));

// workload
const wt=document.getElementById('wtab');
if(DATA.summary.length){const cols=Object.keys(DATA.summary[0]);
 wt.innerHTML='<tr>'+cols.map(c=>`<th class="${/hrs|#/.test(c)?'num':''}">${c}</th>`).join('')+'</tr>'+
 DATA.summary.map(r=>'<tr>'+cols.map(c=>`<td class="${/hrs|#/.test(c)?'num':''}">${r[c]}</td>`).join('')+'</tr>').join('');}

// kpis
const rc=DATA.reconciliation;
const kpiDefs=[['Teachers',rc['Distinct teachers'],''],['Teacher-sessions',rc['Teacher-session assignments (after adjustments)'],''],
 ['Clashes',rc['Clashes detected'],(rc['Clashes detected']==0?'good':'warn')],
 ['Adjustments',rc['Manual adjustments applied'],''],['Needs a teacher',rc['Sessions with NO teacher (excluded)'],'warn']];
document.getElementById('kpis').innerHTML=kpiDefs.map(([k,v,c])=>`<div class="kpi ${c}"><div class="v">${v}</div><div class="k">${k}</div></div>`).join('');

// clashes
const cl=document.getElementById('clashes');
cl.innerHTML=DATA.clashes.length? '<table><tr><th>Teacher</th><th>Day</th><th>Times</th><th>Weeks</th><th>Class A</th><th>Class B</th></tr>'+
 DATA.clashes.map(c=>`<tr><td>${c.Teacher}</td><td>${c.Day}</td><td>${c.Times}</td><td>${c['Overlap weeks']}</td><td>${c['Class A']}</td><td>${c['Class B']}</td></tr>`).join('')+'</table>'
 : '<div class="callout" style="border-color:#bfe6d4;background:#f0faf5;color:#0d7d5a">No clashes remaining - all resolved via adjustments.</div>';

// unassigned
const un=document.getElementById('unassigned');
if(DATA.unassigned.length){const cols=Object.keys(DATA.unassigned[0]);
 un.innerHTML='<tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr>'+
 DATA.unassigned.map(r=>'<tr>'+cols.map(c=>`<td>${r[c]}</td>`).join('')+'</tr>').join('');}

// audit
const badge={'ADJ-APPLIED':['adj','Adjustment'],'COMBINED-MERGE':['merge','Merged'],'NO-TEACHER':['no','No teacher'],
 'NAME-MERGE':['name','Name match'],'AMBIGUOUS-NAME':['name','Ambiguous'],'ADJ-WARN':['no','Warning'],'HOURS?':['no','Hours?']};
document.getElementById('auditlog').innerHTML=DATA.audit.map(a=>{const b=badge[a.type]||['adj',a.type];
 return `<div class="auditrow"><span class="badge ${b[0]}">${b[1]}</span><span>${a.detail}</span></div>`;}).join('');
</script></body></html>"""

if __name__ == "__main__":
    main()
