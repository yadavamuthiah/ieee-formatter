"""
IEEE conference-style formatter for .docx research papers.

Usage (in Jupyter or a terminal):
    from ieee_formatter import format_ieee
    format_ieee("Research paper.docx", "Research paper_IEEE.docx")

Pipeline:  1) classify every paragraph  2) apply IEEE rules  3) save new .docx
"""
import copy
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# --------------------------------------------------------------------------
# IEEE template settings (edit these to support another template)
# --------------------------------------------------------------------------
TEMPLATE = {
    "font": "Times New Roman",
    "page_w": 8.5, "page_h": 11.0,                 # US Letter, inches
    "top": 0.75, "bottom": 1.0, "left": 0.625, "right": 0.625,
    "col_gap_twips": 360,                          # 0.25 inch
    "col_width_in": 3.5,
    "title_pt": 24, "author_pt": 10, "abstract_pt": 9,
    "body_pt": 10, "caption_pt": 8, "ref_pt": 8, "table_pt": 8,
    "first_line_indent_in": 0.2,
}

KEEP_FONTS = {"symbol", "cambria math", "wingdings", "wingdings 2",
              "wingdings 3", "mt extra"}
LEVEL1_STARTS = ("introduction", "literature", "related work", "background",
                 "method", "experiment", "result", "discussion", "conclusion",
                 "future work", "acknowledg", "references")


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def roman(n):
    vals = [(10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def all_runs(p):
    return [Run(r, p) for r in p._p.xpath(".//w:r[not(ancestor::m:oMath)]")]


def has_drawing(p):
    return bool(p._p.xpath(".//w:drawing | .//w:pict | .//w:object"))


def has_math(p):
    return bool(p._p.xpath(".//m:oMath | .//m:oMathPara"))


def get_numpr(p):
    pPr = p._p.pPr
    return pPr.find(qn("w:numPr")) if pPr is not None else None


def remove_numpr(p):
    n = get_numpr(p)
    if n is not None:
        p._p.pPr.remove(n)


def ilvl_of(p):
    n = get_numpr(p)
    if n is None:
        return None
    il = n.find(qn("w:ilvl"))
    return int(il.get(qn("w:val"))) if il is not None else 0


def prepend_run(p, text):
    r = OxmlElement("w:r")
    first = p._p.xpath("./w:r | ./w:hyperlink")
    if first:
        first[0].addprevious(r)
    else:
        p._p.append(r)
    run = Run(r, p)
    run.text = text
    return run


def strip_prefix(p, pattern):
    text = "".join(r.text for r in all_runs(p))
    m = re.match(pattern, text, re.I)
    if not m:
        return
    n = len(m.group(0))
    for r in all_runs(p):
        t = r.text
        if n <= 0:
            break
        if not t:
            continue
        if len(t) <= n:
            r.text = ""
            n -= len(t)
        else:
            r.text = t[n:]
            n = 0


def style_runs(p, size, bold=None, italic=None, small_caps=None):
    for r in all_runs(p):
        rPr = r._r.get_or_add_rPr()
        rf = rPr.find(qn("w:rFonts"))
        current = (rf.get(qn("w:ascii")) or "").lower() if rf is not None else ""
        r.font.size = Pt(size)
        if current in KEEP_FONTS:
            continue
        r.font.name = TEMPLATE["font"]
        rf = rPr.find(qn("w:rFonts"))
        for a in ("w:eastAsia", "w:cs"):
            rf.set(qn(a), TEMPLATE["font"])
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
        if bold is not None:
            r.font.bold = bold
        if italic is not None:
            r.font.italic = italic
        if small_caps is not None:
            r.font.small_caps = small_caps


def para_fmt(p, align=None, first=0.0, left=0.0, before=0, after=0,
             keep_next=False, keep_together=False):
    pf = p.paragraph_format
    if align is not None:
        pf.alignment = align
    pf.first_line_indent = Inches(first)
    pf.left_indent = Inches(left)
    pf.right_indent = Inches(0)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.0
    pf.keep_with_next = keep_next
    pf.keep_together = keep_together


def set_cols(sectPr, num, gap):
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        later = ["w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
                 "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid"]
        anchor = None
        for tag in later:
            anchor = sectPr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(cols)
        else:
            sectPr.append(cols)
    for k in list(cols.attrib):
        del cols.attrib[k]
    for child in list(cols):
        cols.remove(child)
    cols.set(qn("w:num"), str(num))
    cols.set(qn("w:space"), str(gap))


# --------------------------------------------------------------------------
# Step 1: structure detection
# --------------------------------------------------------------------------
# How detection works:
#   The paper is read paragraph by paragraph. Position and text patterns decide
#   the type: the first line is the title, text before 'Abstract' is the author
#   block, text after 'References' is the reference list, 'Fig. N' lines are
#   captions, short Word-numbered lines are headings, and so on.
def classify_all(doc):
    items, state = [], {"section": "front"}
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        low = t.lower()
        kind, level = "body", None
        if not t:
            kind = "figure" if (has_drawing(p) or has_math(p)) else "empty"
            if kind == "figure" and has_math(p) and not has_drawing(p):
                kind = "equation"
        elif i == 0:
            kind = "title"
        elif low.startswith("abstract"):
            state["section"] = "abstract"
            kind = "abstract_heading"
        elif low.startswith("keywords") or low.startswith("index terms"):
            state["section"] = "body"
            kind = "keywords"
        elif low.startswith("references"):
            state["section"] = "references"
            kind = "references_heading"
        elif state["section"] == "front":
            kind = "author_block"
        elif state["section"] == "abstract":
            kind = "abstract_body"
        elif state["section"] == "references":
            kind = "reference"
        elif re.match(r"^(fig\.?|figure|table)\s*\d+", low):
            kind = "caption"
        elif has_math(p) or re.search(r"\(\d+\)\s*$", t):
            kind = "equation"
        elif p.style.name == "List Paragraph" and len(t) < 90:
            kind = "heading"
        elif re.match(r"^\d+(\.\d+)+\s+\S", t) and len(t) < 90:
            kind = "heading"
            level = 2
        elif (get_numpr(p) is not None and len(t) < 120
              and all_runs(p) and all_runs(p)[0].bold):
            kind = "heading"          # bold, auto-numbered short line
            level = 2
        if kind == "heading" and level is None:
            if low.startswith(LEVEL1_STARTS):
                level = 1
            else:
                level = 3 if (ilvl_of(p) or 0) >= 2 else 2
        if kind == "references_heading":
            level = 1
        items.append({"p": p, "kind": kind, "level": level})
    return items


# --------------------------------------------------------------------------
# Step 2: apply IEEE rules
# --------------------------------------------------------------------------
# Main function: reads the file 'src', applies the IEEE rules, saves to 'dst'.
# 'src' and 'dst' can be file names or in-memory files (the website uses in-memory).
def format_ieee(src, dst, verbose=True):
    T = TEMPLATE
    doc = Document(src)
    report = []

    items = classify_all(doc)

    # Normal style baseline
    st = doc.styles["Normal"]
    st.font.name = T["font"]
    st.font.size = Pt(T["body_pt"])

    # ---- page setup -------------------------------------------------------
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(T["page_w"]), Inches(T["page_h"])
    sec.top_margin, sec.bottom_margin = Inches(T["top"]), Inches(T["bottom"])
    sec.left_margin, sec.right_margin = Inches(T["left"]), Inches(T["right"])

    counters = [0, 0, 0]
    ref_no = 0
    to_delete = []
    last_front = None
    abstract_done = False

    for it in items:
        p, kind, level = it["p"], it["kind"], it["level"]

        # ---- Apply the IEEE rule that matches each detected block type ----
        # empty: blank lines are removed (IEEE has no blank-line spacing)
        if kind == "empty":
            to_delete.append(p)

        # title: 24 pt, centered, not bold
        elif kind == "title":
            para_fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=12)
            style_runs(p, T["title_pt"], bold=False)

        # authors / affiliations / emails: centered, full width above the columns
        elif kind == "author_block":
            size = 11 if p is doc.paragraphs[1] else T["author_pt"]
            para_fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=2)
            style_runs(p, size)
            last_front = p

        # the separate 'Abstract' line is deleted; its label is merged into the abstract text
        elif kind == "abstract_heading":
            to_delete.append(p)

        # abstract: starts with 'Abstract—', bold 9 pt, justified
        elif kind == "abstract_body":
            strip_prefix(p, r"^\s*Abstract\s*[:\-–—]?\s*")
            lab = prepend_run(p, "Abstract\u2014")
            para_fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, first=0, after=6)
            style_runs(p, T["abstract_pt"], bold=True)
            lab.font.italic = True
            abstract_done = True

        # keywords: 'Keywords:' becomes 'Index Terms—', bold 9 pt
        elif kind == "keywords":
            strip_prefix(p, r"^\s*(Keywords?|Index Terms)\s*[:\-–—]?\s*")
            lab = prepend_run(p, "Index Terms\u2014")
            para_fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, after=10)
            style_runs(p, T["abstract_pt"], bold=True)
            lab.font.italic = True

        # headings: remove Word's own numbering, then number them ourselves
        #   level 1 -> I., II., III. (centered, small caps)
        #   level 2 -> A., B., C.    (left, italic)
        #   level 3 -> 1), 2), 3)    (indented, italic)
        elif kind in ("heading", "references_heading"):
            remove_numpr(p)
            strip_prefix(p, r"^\s*\d+(\.\d+)*[.)]?\s+")
            for r in all_runs(p):  # ALL CAPS -> Title Case before small caps
                if r.text and r.text.strip().isupper():
                    r.text = r.text.title()
            label = ""
            if kind == "references_heading":
                pass
            elif level == 1:
                counters[0] += 1
                counters[1] = counters[2] = 0
                label = roman(counters[0]) + ". "
            elif level == 2:
                counters[1] += 1
                counters[2] = 0
                label = chr(64 + counters[1]) + ". "
            else:
                counters[2] += 1
                label = f"{counters[2]})  "
            if label:
                prepend_run(p, label)
            if level == 1:
                para_fmt(p, WD_ALIGN_PARAGRAPH.CENTER, before=10, after=4,
                         keep_next=True)
                style_runs(p, T["body_pt"], bold=False, italic=False,
                           small_caps=True)
            elif level == 2:
                para_fmt(p, WD_ALIGN_PARAGRAPH.LEFT, before=6, after=3,
                         keep_next=True)
                style_runs(p, T["body_pt"], bold=False, italic=True,
                           small_caps=False)
            else:
                para_fmt(p, WD_ALIGN_PARAGRAPH.LEFT, left=T["first_line_indent_in"],
                         before=4, after=2, keep_next=True)
                style_runs(p, T["body_pt"], bold=False, italic=True,
                           small_caps=False)

        # figure: the picture's paragraph is centered and kept with its caption
        elif kind == "figure":
            para_fmt(p, WD_ALIGN_PARAGRAPH.CENTER, before=4, after=2,
                     keep_next=True)

        # caption: rewritten as 'Fig. N.' in 8 pt, centered
        elif kind == "caption":
            m = re.match(r"^\s*(fig\.?|figure)\s*(\d+)", p.text, re.I)
            if m:
                strip_prefix(p, r"^\s*(fig\.?|figure)\s*\d+\s*[.:\-–]?\s*")
                prepend_run(p, f"Fig. {m.group(2)}. ")
            para_fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=8, keep_together=True)
            style_runs(p, T["caption_pt"], bold=False)

        # equation: extra tabs removed, equation number pushed to the right edge
        elif kind == "equation":
            # collapse runs of tabs so the equation number sits at the right edge
            tabs = p._p.xpath("./w:r/w:tab")
            for tb in tabs[:-1]:
                tb.getparent().remove(tb)
            para_fmt(p, WD_ALIGN_PARAGRAPH.LEFT, before=3, after=3)
            p.paragraph_format.tab_stops.clear_all()
            p.paragraph_format.tab_stops.add_tab_stop(
                Inches(T["col_width_in"]), WD_TAB_ALIGNMENT.RIGHT)
            style_runs(p, T["body_pt"])

        # reference: renumbered [1], [2], ... in 8 pt with a hanging indent
        elif kind == "reference":
            ref_no += 1
            remove_numpr(p)
            strip_prefix(p, r"^\s*\[?\d+\]\s*")
            prepend_run(p, f"[{ref_no}]\t")
            para_fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, first=-0.3, left=0.3, after=2)
            p.paragraph_format.tab_stops.clear_all()
            p.paragraph_format.tab_stops.add_tab_stop(Inches(0.3))
            style_runs(p, T["ref_pt"])

        # everything else is normal body text: 10 pt, justified, small first-line indent
        else:  # body text
            if has_drawing(p):
                para_fmt(p, WD_ALIGN_PARAGRAPH.CENTER, keep_next=True)
            elif p.text.strip().startswith("="):
                para_fmt(p, WD_ALIGN_PARAGRAPH.LEFT, left=0.2)
                style_runs(p, T["body_pt"])
            else:
                para_fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first=T["first_line_indent_in"])
                style_runs(p, T["body_pt"])

    for p in to_delete:
        p._p.getparent().remove(p._p)

    # (images wider than one column are shrunk so they do not overflow)
    # ---- scale images to column width ----------------------------------
    max_w = int(3.4 * 914400)
    scaled = anchored = 0
    for d in doc.element.body.xpath(".//w:drawing"):
        if d.xpath("./wp:anchor"):
            anchored += 1
        ext = d.xpath(".//wp:extent")
        if not ext:
            continue
        cx, cy = int(ext[0].get("cx")), int(ext[0].get("cy"))
        if cx > max_w:
            ratio = max_w / cx
            ext[0].set("cx", str(int(cx * ratio)))
            ext[0].set("cy", str(int(cy * ratio)))
            for a in d.xpath(".//a:xfrm/a:ext"):
                a.set("cx", str(int(int(a.get("cx")) * ratio)))
                a.set("cy", str(int(int(a.get("cy")) * ratio)))
            scaled += 1
    report.append(f"Images resized to fit column: {scaled}")
    if anchored:
        report.append(f"Floating (anchored) images: {anchored} - check these in Word")

    # (table widths are reduced to the column width and text set to 8 pt)
    # ---- tables: shrink to column width, 8 pt font ----------------------
    max_tw = int(3.45 * 1440)
    for t in doc.tables:
        grid = t._tbl.xpath("./w:tblGrid/w:gridCol")
        total = sum(int(g.get(qn("w:w"))) for g in grid) or 1
        if total > max_tw:
            f = max_tw / total
            for g in grid:
                g.set(qn("w:w"), str(int(int(g.get(qn("w:w"))) * f)))
            for tcW in t._tbl.xpath(".//w:tcW"):
                if tcW.get(qn("w:type")) == "dxa":
                    tcW.set(qn("w:w"), str(int(int(tcW.get(qn("w:w"))) * f)))
        tblW = t._tbl.tblPr.find(qn("w:tblW"))
        if tblW is not None:
            tblW.set(qn("w:type"), "dxa")
            tblW.set(qn("w:w"), str(min(total, max_tw)))
        for row in t.rows:
            for cell in row.cells:
                for cp in cell.paragraphs:
                    para_fmt(cp, WD_ALIGN_PARAGRAPH.CENTER)
                    style_runs(cp, T["table_pt"])
    report.append(f"Tables formatted: {len(doc.tables)}")

    # (a continuous section break keeps the title full width; the rest uses 2 columns)
    # ---- two columns: title block full width, rest in 2 columns ---------
    sectPr = sec._sectPr
    if last_front is not None:
        front = copy.deepcopy(sectPr)
        set_cols(front, 1, T["col_gap_twips"])
        last_front._p.get_or_add_pPr().append(front)
        last_front.paragraph_format.space_after = Pt(12)
    set_cols(sectPr, 2, T["col_gap_twips"])
    sec.start_type = WD_SECTION.CONTINUOUS

    doc.save(dst)

    summary = {}
    for it in items:
        summary[it["kind"]] = summary.get(it["kind"], 0) + 1
    report.append("Detected blocks: " + ", ".join(f"{k}={v}" for k, v in summary.items()))
    report.append(f"Headings: {counters[0]} main, references renumbered: {ref_no}")
    if verbose:
        print("\n".join(report))
    return report


if __name__ == "__main__":
    import sys
    src = sys.argv[1] if len(sys.argv) > 1 else "input.docx"
    dst = sys.argv[2] if len(sys.argv) > 2 else "output_IEEE.docx"
    format_ieee(src, dst)
