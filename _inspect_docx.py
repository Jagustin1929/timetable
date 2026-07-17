"""Dump the ordered structure of a .docx (paragraphs w/ style + tables) using stdlib only."""
import sys, zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def text_of(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))

def para_style(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None, False
    st = ppr.find(W + "pStyle")
    style = st.get(W + "val") if st is not None else None
    bold = ppr.find(W + "rPr/" + W + "b") is not None
    return style, bold

def cell_text(tc):
    return " ".join(text_of(p).strip() for p in tc.findall(W + "p")).strip()

def dump(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(W + "body")
    print("=" * 90)
    print("FILE:", path.split("/")[-1])
    print("=" * 90)
    tbl_n = 0
    for child in body:
        tag = child.tag.replace(W, "")
        if tag == "p":
            txt = text_of(child).strip()
            style, bold = para_style(child)
            if txt or style:
                flag = f"[style={style}]" if style else ""
                flag += "[BOLD]" if bold else ""
                print(f"P {flag} {txt!r}")
        elif tag == "tbl":
            tbl_n += 1
            rows = child.findall(W + "tr")
            print(f"\n--- TABLE {tbl_n}: {len(rows)} rows ---")
            for ri, tr in enumerate(rows):
                cells = tr.findall(W + "tc")
                vals = [cell_text(tc) for tc in cells]
                shown = " | ".join(v.replace("\n", " / ") for v in vals)
                if len(shown) > 260:
                    shown = shown[:260] + " …"
                print(f"  R{ri} ({len(cells)}c): {shown}")
            print("--- end table ---\n")

if __name__ == "__main__":
    dump(sys.argv[1])
