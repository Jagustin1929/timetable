"""Minimal, dependency-free PDF writer (standard 14 fonts, no embedding).
Enough to lay out simple tabular teacher timetables. Coordinates are given with
the ORIGIN AT TOP-LEFT (y grows downward); we convert to PDF's bottom-left origin."""


class SimplePDF:
    def __init__(self, size=(842, 595)):      # A4 landscape in points
        self.w, self.h = size
        self.pages = []

    def add_page(self):
        self.pages.append([])
        return self

    # --- text helpers ---
    @staticmethod
    def _san(s):
        s = str(s)
        for k, v in {"\u2013": "-", "\u2014": "-", "\u2018": "'", "\u2019": "'",
                     "\u201c": '"', "\u201d": '"', "\u2026": "...", "\u00a0": " "}.items():
            s = s.replace(k, v)
        return s.encode("latin-1", "replace").decode("latin-1")

    @staticmethod
    def _esc(s):
        return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    def text(self, x, y, size, s, bold=False):
        f = "F2" if bold else "F1"
        self.pages[-1].append(
            f"BT /{f} {size:.2f} Tf 1 0 0 1 {x:.2f} {self.h - y:.2f} Tm "
            f"({self._esc(self._san(s))}) Tj ET")

    def line(self, x1, y1, x2, y2, w=0.5, gray=0.0):
        self.pages[-1].append(
            f"{gray:.2f} G {w:.2f} w {x1:.2f} {self.h - y1:.2f} m "
            f"{x2:.2f} {self.h - y2:.2f} l S 0 G")

    def fill_rect(self, x, y, w, h, gray=0.9):
        self.pages[-1].append(
            f"{gray:.2f} g {x:.2f} {self.h - y - h:.2f} {w:.2f} {h:.2f} re f 0 g")

    # --- serialisation ---
    def output(self):
        objects = {
            1: b"<< /Type /Catalog /Pages 2 0 R >>",
            3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            4: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
        }
        kids, num = [], 5
        for ops in self.pages:
            content = ("\n".join(ops)).encode("latin-1", "replace")
            cnum, pnum = num, num + 1
            num += 2
            objects[cnum] = b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream"
            objects[pnum] = (
                f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {self.w} {self.h}] "
                f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                f"/Contents {cnum} 0 R >>").encode("latin-1")
            kids.append(f"{pnum} 0 R")
        objects[2] = (f"<< /Type /Pages /Kids [{' '.join(kids)}] "
                      f"/Count {len(self.pages)} >>").encode("latin-1")

        out = bytearray(b"%PDF-1.4\n")
        maxn = max(objects)
        offsets = {}
        for n in range(1, maxn + 1):
            if n not in objects:
                continue
            offsets[n] = len(out)
            out += f"{n} 0 obj\n".encode("latin-1") + objects[n] + b"\nendobj\n"
        xref = len(out)
        out += f"xref\n0 {maxn + 1}\n".encode("latin-1")
        out += b"0000000000 65535 f \n"
        for n in range(1, maxn + 1):
            out += (f"{offsets[n]:010d} 00000 n \n" if n in offsets
                    else "0000000000 65535 f \n").encode("latin-1")
        out += (f"trailer\n<< /Size {maxn + 1} /Root 1 0 R >>\n"
                f"startxref\n{xref}\n%%EOF").encode("latin-1")
        return bytes(out)
