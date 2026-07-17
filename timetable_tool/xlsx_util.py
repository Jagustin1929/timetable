"""Shared workbook writer: uses openpyxl if available, else a stdlib zip/XML writer.
   write_workbook(path, sheets) where sheets = [(sheet_name, header_list, rows_list_of_lists)]."""
import re, zipfile


def _esc(v):
    s = "" if v is None else str(v)
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def _col(n):
    s = ""
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s


def write_workbook(path, sheets):
    try:
        import openpyxl
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        used = set()
        for name, header, rows in sheets:
            t = _safe_title(name, used)
            ws = wb.create_sheet(t); ws.append(list(header))
            for row in rows:
                ws.append(list(row))
        wb.save(path)
        return "openpyxl"
    except ImportError:
        pass

    def sheet_xml(header, rows):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
        for ri, row in enumerate([header] + list(rows), start=1):
            out.append(f'<row r="{ri}">')
            for ci, val in enumerate(row):
                out.append(f'<c r="{_col(ci)}{ri}" t="inlineStr"><is>'
                           f'<t xml:space="preserve">{_esc(val)}</t></is></c>')
            out.append('</row>')
        out.append('</sheetData></worksheet>')
        return "".join(out)

    used = set()
    named = [(_safe_title(n, used), h, r) for n, h, r in sheets]

    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(len(named)):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" '
                  f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append('</Types>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb_sheets, wb_rels = [], []
    for i, (name, _, _) in enumerate(named):
        sid = i + 1
        wb_sheets.append(f'<sheet name="{_esc(name)}" sheetId="{sid}" r:id="rId{sid}"/>')
        wb_rels.append(f'<Relationship Id="rId{sid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sid}.xml"/>')
    workbook_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                    + "".join(wb_sheets) + '</sheets></workbook>')
    workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     + "".join(wb_rels) + '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        for i, (_, header, rows) in enumerate(named):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", sheet_xml(header, rows))
    return "stdlib"


def _safe_title(name, used):
    t = re.sub(r"[\\/*?:\[\]]", " ", str(name))[:31].strip() or "Sheet"
    base, i = t, 1
    while t.lower() in used:
        suf = f"_{i}"; t = base[:31 - len(suf)] + suf; i += 1
    used.add(t.lower())
    return t
