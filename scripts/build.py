#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""english-resume builder — YAML content in, one-page A4 resume out (DOCX + PDF).

Usage
-----
    python scripts/build.py examples/haiyang-zhao.yaml
    python scripts/build.py resume.yaml -o out --max-pages 1 --png

Why the PDF is drawn by hand with PyMuPDF instead of Word/Chrome:
Word COM automation and headless-browser printing are both blocked in locked-down
environments. Drawing the page directly makes the result reproducible and lets the
builder measure every line, so the auto-shrink loop can guarantee one page.

Layout is fixed (see references/layout-spec.md); content lives in YAML.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import yaml

CM = 28.3464567  # 1 cm in PDF points

# --------------------------------------------------------------------------- #
# Geometry / typography — the part that must not drift
# --------------------------------------------------------------------------- #
SPEC = {
    "page_w_cm": 21.0,
    "page_h_cm": 29.7,
    "margin_t_cm": 1.70,
    "margin_b_cm": 1.60,
    "margin_l_cm": 1.90,
    "margin_r_cm": 1.90,
    "name_pt": 16.0,
    "header_pt": 11.5,
    "body_pt": 10.0,
    "glyph_ratio": 0.80,   # bullet dot size relative to body
    "line_factor": 1.16,   # PDF baseline advance per body line
    "gap_name_after": 2.0,
    "gap_contact_after": 2.0,
    "header_before": 8.0,
    "header_after": 3.0,
    "entry_before": 2.0,
    "entry_after": 1.0,
    "bullet_after": 1.0,
    "col_date_cm": 3.35,   # date column width (left tab)
    "col_org_cm": 8.55,    # org column ends / role column starts
    "border_sz": 6,        # docx bottom-border weight (1/8 pt)
    "border_pt": 0.75,     # pdf rule weight
}

FONT_FILES = {
    "r": ("times.ttf", "tiro"),
    "b": ("timesbd.ttf", "tibo"),
    "i": ("timesi.ttf", "tiit"),
    "bi": ("timesbi.ttf", "tibi"),
}
_WIN_FONTS = Path("C:/Windows/Fonts")


def _resolve_fontfs(kind: str) -> str | None:
    fname, _ = FONT_FILES[kind]
    p = _WIN_FONTS / fname
    return str(p) if p.exists() else None


def _measure(kind: str):
    import pymupdf

    fpath = _resolve_fontfs(kind)
    if fpath:
        return pymupdf.Font(fontfile=fpath)
    return pymupdf.Font(fontname=FONT_FILES[kind][1])


class Metrics:
    """Text width in points, using the real Times New Roman metrics."""

    def __init__(self):
        self._cache = {}

    def font(self, bold: bool, italic: bool):
        key = "bi" if (bold and italic) else "b" if bold else "i" if italic else "r"
        if key not in self._cache:
            self._cache[key] = _measure(key)
        return self._cache[key]

    def width(self, text: str, size: float, bold: bool = False, italic: bool = False) -> float:
        return self.font(bold, italic).text_length(text or "", fontsize=size)


M = Metrics()


# --------------------------------------------------------------------------- #
# Content model
# --------------------------------------------------------------------------- #
def load_content(path: Path) -> dict:
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"{path}: top level must be a mapping")
    if not raw.get("name"):
        raise ValueError(f"{path}: 'name' is required")
    for i, sec in enumerate(raw.get("sections") or []):
        if "type" not in sec:
            raise ValueError(f"{path}: sections[{i}] needs a 'type'")
        if sec["type"] not in ("paragraph", "entries", "skills"):
            raise ValueError(f"{path}: sections[{i}] unknown type {sec['type']!r}")
    return raw


def wrap_text(text, size, width_pt, bold=False, italic=False):
    """Greedy word wrap measured with real font metrics."""
    out, cur = [], ""
    for word in str(text).split():
        trial = word if not cur else cur + " " + word
        if M.width(trial, size, bold, italic) <= width_pt:
            cur = trial
        else:
            if cur:
                out.append(cur)
            cur = word
    if cur:
        out.append(cur)
    return out


# --------------------------------------------------------------------------- #
# DOCX
# --------------------------------------------------------------------------- #
def _setup_normal(doc, body_pt):
    from docx.shared import Pt
    from docx.oxml.ns import qn

    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(body_pt)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        from docx.oxml import OxmlElement

        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), "Times New Roman")
    pf = st.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def _run(p, text, size, bold=False, italic=False):
    from docx.shared import Pt

    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    return r


def _para_space(p, before=None, after=None):
    from docx.shared import Pt

    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)


def _bottom_border(p, sz=6):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pbdr.append(bottom)
    pPr.append(pbdr)


def build_docx(content, spec, out_path: Path):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.shared import Cm, Pt

    body = spec["body_pt"]
    header = spec["header_pt"]
    glyph = round(body * spec["glyph_ratio"], 1)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(spec["page_w_cm"]), Cm(spec["page_h_cm"])
    sec.top_margin, sec.bottom_margin = Cm(spec["margin_t_cm"]), Cm(spec["margin_b_cm"])
    sec.left_margin, sec.right_margin = Cm(spec["margin_l_cm"]), Cm(spec["margin_r_cm"])
    _setup_normal(doc, body)

    content_w_cm = spec["page_w_cm"] - spec["margin_l_cm"] - spec["margin_r_cm"]
    text_w_cm = content_w_cm

    # --- header -------------------------------------------------------- #
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _para_space(p, after=spec["gap_name_after"])
    _run(p, content["name"], spec["name_pt"], bold=True)

    if content.get("contact"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _para_space(p, after=spec["gap_contact_after"])
        _run(p, content["contact"], body)

    if content.get("availability"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _para_space(p, after=spec["gap_contact_after"])
        _run(p, content["availability"], body, bold=True)

    # --- sections ------------------------------------------------------- #
    for sect in content.get("sections") or []:
        title = sect.get("title")
        if title:
            p = doc.add_paragraph()
            _para_space(p, before=spec["header_before"], after=spec["header_after"])
            _run(p, title.upper(), header, bold=True)
            _bottom_border(p, spec["border_sz"])

        if sect["type"] == "paragraph":
            for para in _as_list(sect.get("text")):
                p = doc.add_paragraph()
                _para_space(p, after=spec["bullet_after"])
                _run(p, para, body)

        elif sect["type"] == "entries":
            for e in sect.get("entries") or []:
                _docx_entry(doc, e, spec, body, glyph, text_w_cm)

        elif sect["type"] == "skills":
            for g in sect.get("groups") or []:
                if g.get("label"):
                    p = doc.add_paragraph()
                    _para_space(p, after=spec["bullet_after"])
                    _run(p, g["label"], body, bold=True)
                for item in _as_list(g.get("items")):
                    _docx_bullet(doc, item, spec, body, glyph)

    core = doc.core_properties
    core.author = content["name"].title()
    core.title = f"{content['name']} — Resume"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def _as_list(v):
    if v is None:
        return []
    return v if isinstance(v, list) else [v]


def _fits_one_line(e, spec, body) -> bool:
    """Can date | org | role | location share one line at this body size?"""
    avail = (spec["page_w_cm"] - spec["margin_l_cm"] - spec["margin_r_cm"]) * CM
    x_org = spec["col_date_cm"] * CM
    x_role = spec["col_org_cm"] * CM
    # the org string must stop before the role column starts
    if x_org + M.width(e.get("org", ""), body, bold=True) + 6 > x_role:
        return False
    need = M.width(e.get("role", ""), body, bold=True) + M.width(
        e.get("location", ""), body, bold=True
    )
    return x_role + need + 8 <= avail


def _docx_entry(doc, e, spec, body, glyph, text_w_cm):
    from docx.enum.text import WD_TAB_ALIGNMENT
    from docx.shared import Cm

    one_line = _fits_one_line(e, spec, body)
    date = e.get("date", "")
    org = e.get("org", "")
    role = e.get("role", "")
    loc = e.get("location", "")

    p = doc.add_paragraph()
    _para_space(p, before=spec["entry_before"], after=spec["entry_after"])
    ts = p.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(spec["col_date_cm"]), WD_TAB_ALIGNMENT.LEFT)
    ts.add_tab_stop(Cm(spec["col_org_cm"]), WD_TAB_ALIGNMENT.LEFT)
    ts.add_tab_stop(Cm(text_w_cm), WD_TAB_ALIGNMENT.RIGHT)

    _run(p, date, body, bold=True)
    if org:
        _run(p, f"\t{org}", body, bold=True)
    if one_line:
        if role:
            _run(p, f"\t{role}", body, bold=True)
        if loc:
            _run(p, f"\t{loc}", body, bold=True)
    else:
        if loc:
            _run(p, f"\t\t{loc}", body, bold=True)
        if role:
            p2 = doc.add_paragraph()
            _para_space(p2, after=spec["entry_after"])
            _run(p2, role, body, bold=False, italic=True)
            if loc:
                p2.paragraph_format.tab_stops.add_tab_stop(Cm(text_w_cm), WD_TAB_ALIGNMENT.RIGHT)
                _run(p2, f"\t{loc}", body, bold=True)

    for b in _as_list(e.get("bullets")):
        _docx_bullet(doc, b, spec, body, glyph)


def _docx_bullet(doc, text, spec, body, glyph):
    from docx.enum.text import WD_TAB_ALIGNMENT
    from docx.shared import Cm

    p = doc.add_paragraph()
    _para_space(p, after=spec["bullet_after"])
    pf = p.paragraph_format
    pf.left_indent = Cm(0.55)
    pf.first_line_indent = Cm(-0.45)
    pf.tab_stops.add_tab_stop(Cm(0.55), WD_TAB_ALIGNMENT.LEFT)
    _run(p, "\u25cf\t", glyph)
    _run(p, text, body)


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #
BASE14 = {"r": "tiro", "b": "tibo", "i": "tiit", "bi": "tibi"}


class PdfRenderer:
    """Draws the page directly, so no Word and no headless browser is needed."""

    def __init__(self, content: dict, spec: dict):
        import pymupdf

        self.pymupdf = pymupdf
        self.content = content
        self.spec = spec
        self.pw = spec["page_w_cm"] * CM
        self.ph = spec["page_h_cm"] * CM
        self.ml = spec["margin_l_cm"] * CM
        self.mr = spec["margin_r_cm"] * CM
        self.mt = spec["margin_t_cm"] * CM
        self.mb = spec["margin_b_cm"] * CM
        self.avail = self.pw - self.ml - self.mr
        self.body = spec["body_pt"]
        self.header = spec["header_pt"]
        self.glyph = self.body * spec["glyph_ratio"]
        self.lh = spec["line_factor"]
        self.doc = pymupdf.open()
        self.page = None
        self.fn = {}
        self.top = self.mt
        self._new_page()

    # -- page --------------------------------------------------------- #
    def _new_page(self):
        self.page = self.doc.new_page(width=self.pw, height=self.ph)
        self.fn = {}
        for kind in ("r", "b", "i", "bi"):
            fpath = _resolve_fontfs(kind)
            if fpath:
                self.page.insert_font(fontname=kind, fontfile=fpath)
                self.fn[kind] = kind
            else:
                self.fn[kind] = BASE14[kind]
        self.top = self.mt

    def _need(self, h):
        if self.top + h > self.ph - self.mb:
            self._new_page()

    # -- primitives --------------------------------------------------- #
    def _text(self, x, baseline, text, kind, size):
        self.page.insert_text(
            (x, baseline), text, fontname=self.fn[kind], fontsize=size, color=(0, 0, 0)
        )

    def _put(self, baseline, text, kind, size, align="left", x=None):
        w = M.width(text, size, kind in ("b", "bi"), kind in ("i", "bi"))
        if align == "center":
            x = (self.pw - w) / 2
        elif align == "right":
            x = self.pw - self.mr - w
        else:
            x = self.ml if x is None else x
        self._text(x, baseline, text, kind, size)
        return w

    def _flow(self, text, size=None, kind="r", align="left", after=0.0):
        size = self.body if size is None else size
        h = size * self.lh
        for ln in wrap_text(text, size, self.avail, kind in ("b", "bi"), kind in ("i", "bi")):
            self._need(h)
            self._put(self.top + (h - size) / 2 + 0.75 * size, ln, kind, size, align=align)
            self.top += h
        self.top += after

    # -- blocks ------------------------------------------------------- #
    def _section_rule(self, title):
        h = self.header * 1.25
        self._need(self.spec["header_before"] + h + self.spec["header_after"])
        self.top += self.spec["header_before"]
        self._put(self.top + (h - self.header) / 2 + 0.75 * self.header,
                  title.upper(), "b", self.header)
        self.top += h
        self.page.draw_line(
            self.pymupdf.Point(self.ml, self.top + 0.3),
            self.pymupdf.Point(self.pw - self.mr, self.top + 0.3),
            color=(0, 0, 0),
            width=self.spec["border_pt"],
        )
        self.top += self.spec["header_after"]

    def _entry(self, e):
        body = self.body
        h = body * self.lh
        date = e.get("date") or ""
        org = e.get("org") or ""
        role = e.get("role") or ""
        loc = e.get("location") or ""
        one_line = _fits_one_line(e, self.spec, body)


        x_date = self.ml
        x_org = self.ml + self.spec["col_date_cm"] * CM
        x_role = self.ml + self.spec["col_org_cm"] * CM

        # A long org string (award name, long project title) must wrap instead of
        # running off the page. When it wraps we stack it under the date line.
        org_lines = (
            wrap_text(org, body, self.pw - self.mr - x_org, bold=True) if org else []
        )
        stacked = len(org_lines) > 1
        if stacked:
            one_line = False

        rows = (1 + len(org_lines)) if stacked else 1
        if (not one_line) and role:
            rows += 1
        self._need(h * rows)
        self.top += self.spec["entry_before"]

        def baseline():
            return self.top + (h - body) / 2 + 0.75 * body

        base = baseline()
        if date:
            self._text(x_date, base, date, "b", body)

        if stacked:
            if loc:
                self._put(base, loc, "b", body, align="right")
            for ln in org_lines:
                self.top += h
                self._need(h)
                self._text(x_org, baseline(), ln, "b", body)
        else:
            if org_lines:
                self._text(x_org, base, org_lines[0], "b", body)
            if one_line:
                if role:
                    self._text(x_role, base, role, "b", body)
                if loc:
                    self._put(base, loc, "b", body, align="right")
            else:
                if loc:
                    self._put(base, loc, "b", body, align="right")
                if role:
                    self.top += h
                    self._need(h)
                    self._text(x_org, baseline(), role, "i", body)

        self.top += h + self.spec["entry_after"]
        for b in _as_list(e.get("bullets")):
            self._bullet(b)


    def _bullet(self, text):
        body = self.body
        h = body * self.lh
        x_dot = self.ml + 0.10 * CM
        x_txt = self.ml + 0.55 * CM
        width = self.avail - 0.55 * CM
        for i, ln in enumerate(wrap_text(text, body, width)):
            self._need(h)
            base = self.top + (h - body) / 2 + 0.75 * body
            if i == 0:
                self.page.insert_text(
                    (x_dot, base), "\u25cf", fontname=self.fn["r"],
                    fontsize=self.glyph, color=(0, 0, 0),
                )
            self._text(x_txt, base, ln, "r", body)
            self.top += h
        self.top += self.spec["bullet_after"]

    # -- assemble ----------------------------------------------------- #
    def draw(self):
        c, spec = self.content, self.spec
        h = spec["name_pt"] * self.lh
        self._put(self.top + (h - spec["name_pt"]) / 2 + 0.75 * spec["name_pt"],
                  c["name"], "b", spec["name_pt"], align="center")
        self.top += h + spec["gap_name_after"]
        if c.get("contact"):
            self._flow(c["contact"], align="center", after=spec["gap_contact_after"])
        if c.get("availability"):
            self._flow(c["availability"], kind="b", align="center",
                       after=spec["gap_contact_after"])
        for sect in c.get("sections") or []:
            if sect.get("title"):
                self._section_rule(sect["title"])
            kind = sect["type"]
            if kind == "paragraph":
                for para in _as_list(sect.get("text")):
                    self._flow(para, after=spec["bullet_after"])
            elif kind == "entries":
                for e in sect.get("entries") or []:
                    self._entry(e)
            elif kind == "skills":
                for g in sect.get("groups") or []:
                    if g.get("label"):
                        self._flow(g["label"], kind="b", after=spec["bullet_after"])
                    for item in _as_list(g.get("items")):
                        self._bullet(item)
        return self.doc


def render_doc(content, spec):
    """Render to an open PyMuPDF document (caller owns closing it)."""
    return PdfRenderer(content, spec).draw()


def save_pdf(doc, out_path: Path, png: bool = False):
    import pymupdf

    doc.set_metadata({"title": "Resume", "author": ""})
    try:
        doc.subset_fonts()
    except Exception:
        pass
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path), garbage=4, deflate=True)
    if png:
        d = pymupdf.open(str(out_path))
        d[0].get_pixmap(dpi=160).save(str(out_path.with_suffix(".png")))
        d.close()


def fit_body_pt(content, spec, max_pages=1, min_pt=8.0):
    """Shrink the body size until the resume fits max_pages. Returns (body_pt, pages)."""
    size = spec["body_pt"]
    pages = 99
    while size >= min_pt:
        doc = render_doc(content, dict(spec, body_pt=round(size, 1)))
        pages = doc.page_count
        doc.close()
        if pages <= max_pages:
            return round(size, 1), pages
        size -= 0.1
    return round(size + 0.1, 1), pages


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="Build a one-page English resume (DOCX + PDF).")
    ap.add_argument("content", type=Path, help="YAML file describing the resume")
    ap.add_argument("-o", "--outdir", type=Path, default=None, help="output directory")
    ap.add_argument("-n", "--stem", default=None, help="output file stem")
    ap.add_argument("--max-pages", type=int, default=1, help="shrink until it fits N pages (0 = off)")
    ap.add_argument("--body-pt", type=float, default=None, help="force a body size, skip auto-fit")
    ap.add_argument("--png", action="store_true", help="also write a first-page PNG preview")
    args = ap.parse_args(argv)

    content = load_content(args.content)
    spec = dict(SPEC)
    outdir = args.outdir or Path("out")
    stem = args.stem or _slug(content["name"]) + "-Resume-EN"

    if args.body_pt:
        spec["body_pt"] = args.body_pt
        doc = render_doc(content, spec)
        pages = doc.page_count
        doc.close()
    elif args.max_pages > 0:
        spec["body_pt"], pages = fit_body_pt(content, spec, args.max_pages)
    else:
        pages = 0

    docx_path = outdir / f"{stem}.docx"
    pdf_path = outdir / f"{stem}.pdf"
    build_docx(content, spec, docx_path)
    doc = render_doc(content, spec)
    save_pdf(doc, pdf_path, png=args.png)
    doc.close()
    if spec["body_pt"] < 9.0:
        print(
            f"HINT: body is {spec['body_pt']} pt, below the comfortable floor (9.0). "
            "Cut a bullet instead of shrinking further.",
            file=sys.stderr,
        )

    print(f"body   : {spec['body_pt']} pt")
    print(f"pages  : {pages}" if pages else "pages  : (not checked)")
    print(f"docx   : {docx_path}")
    print(f"pdf    : {pdf_path}")
    if args.png:
        print(f"png    : {pdf_path.with_suffix('.png')}")
    if args.max_pages and pages > args.max_pages:
        print(
            f"WARN: still {pages} pages at {spec['body_pt']} pt — "
            "cut content rather than shrinking further.",
            file=sys.stderr,
        )
        return 1
    return 0

def _slug(name: str) -> str:
    """Build a file stem from a name written either family-first or family-last."""
    parts = [p.strip(",") for p in name.split() if p.strip(",")]
    if len(parts) < 2:
        return "Resume"
    first, last = parts[0], parts[-1]
    family_first = first.isupper() and not last.isupper()
    return f"{first}-{last}" if family_first else f"{last}-{first}"


if __name__ == "__main__":
    raise SystemExit(main())
