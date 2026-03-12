#!/usr/bin/env python3
"""
Auto-adjust Word document format:
1. Delete "No. of Samples" and "Banner" columns from tables
2. Split "Result/ Rating" into two columns: "Result" and "Rating"
3. Change all fonts to Tahoma, 10pt
4. Change orientation from landscape to portrait, auto-fit tables to window
"""

import copy
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---------------------------------------------------------------------------
# XML-level helpers for table column manipulation
# ---------------------------------------------------------------------------

def get_grid_span(cell_el):
    tc_pr = cell_el.find(f"{W}tcPr")
    if tc_pr is not None:
        gs = tc_pr.find(f"{W}gridSpan")
        if gs is not None:
            return int(gs.get(f"{W}val", "1"))
    return 1


def set_grid_span(cell_el, span):
    tc_pr = cell_el.find(f"{W}tcPr")
    if tc_pr is None:
        tc_pr = etree.SubElement(cell_el, f"{W}tcPr")
        cell_el.insert(0, tc_pr)
    gs = tc_pr.find(f"{W}gridSpan")
    if span <= 1:
        if gs is not None:
            tc_pr.remove(gs)
    else:
        if gs is None:
            gs = etree.SubElement(tc_pr, f"{W}gridSpan")
        gs.set(f"{W}val", str(span))


def cell_text(cell_el):
    return "".join(t.text for t in cell_el.iter(f"{W}t") if t.text).strip()


def find_column_index(tbl_el, match_fn):
    """Return the grid-column index of the first header cell whose text satisfies *match_fn*, or -1."""
    rows = tbl_el.findall(f"{W}tr")
    if not rows:
        return -1
    col = 0
    for cell in rows[0].findall(f"{W}tc"):
        if match_fn(cell_text(cell)):
            return col
        col += get_grid_span(cell)
    return -1


def delete_column(tbl_el, col_idx):
    """Remove the grid column at *col_idx* from the table."""
    tbl_grid = tbl_el.find(f"{W}tblGrid")
    if tbl_grid is not None:
        grid_cols = tbl_grid.findall(f"{W}gridCol")
        if col_idx < len(grid_cols):
            tbl_grid.remove(grid_cols[col_idx])

    for tr in tbl_el.findall(f"{W}tr"):
        pos = 0
        for cell in tr.findall(f"{W}tc"):
            span = get_grid_span(cell)
            if pos <= col_idx < pos + span:
                if span == 1:
                    tr.remove(cell)
                else:
                    set_grid_span(cell, span - 1)
                break
            pos += span


def _clone_cell_shell(cell_el):
    """Create an empty <w:tc> that inherits tcPr (minus gridSpan) from *cell_el*."""
    new = etree.Element(f"{W}tc")
    old_pr = cell_el.find(f"{W}tcPr")
    if old_pr is not None:
        pr = copy.deepcopy(old_pr)
        gs = pr.find(f"{W}gridSpan")
        if gs is not None:
            pr.remove(gs)
        new.append(pr)
    p = etree.SubElement(new, f"{W}p")
    return new


def _copy_run_format(src_cell):
    """Return a deep-copied <w:rPr> from the first run of *src_cell*, or None."""
    for r in src_cell.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            return copy.deepcopy(rpr)
    return None


def split_result_rating(tbl_el):
    """Find the 'Result / Rating' column and split it into two columns."""
    col_idx = find_column_index(
        tbl_el,
        lambda txt: "result" in txt.lower() and "rating" in txt.lower(),
    )
    if col_idx < 0:
        return False

    # Widen tblGrid — each new column gets text+padding width
    result_w = (len("result") + 2) * DXA_PER_CHAR
    rating_w = (len("rating") + 2) * DXA_PER_CHAR
    tbl_grid = tbl_el.find(f"{W}tblGrid")
    if tbl_grid is not None:
        gcols = tbl_grid.findall(f"{W}gridCol")
        if col_idx < len(gcols):
            old = gcols[col_idx]
            old.set(f"{W}w", str(result_w))
            ng = etree.Element(f"{W}gridCol")
            ng.set(f"{W}w", str(rating_w))
            old.addnext(ng)

    rows = tbl_el.findall(f"{W}tr")
    header_rpr = None

    for ri, tr in enumerate(rows):
        pos = 0
        for cell in tr.findall(f"{W}tc"):
            span = get_grid_span(cell)
            if pos <= col_idx < pos + span:
                if span == 1:
                    new_cell = _clone_cell_shell(cell)
                    cell.addnext(new_cell)
                    if ri == 0:
                        header_rpr = _copy_run_format(cell)
                        # Keep only the first paragraph ("Result") in original cell
                        paras = cell.findall(f"{W}p")
                        for p in paras[1:]:
                            cell.remove(p)
                        # Set "Rating" in new cell with matching format
                        p = new_cell.find(f"{W}p")
                        if p is None:
                            p = etree.SubElement(new_cell, f"{W}p")
                        src_ppr = paras[0].find(f"{W}pPr") if paras else None
                        if src_ppr is not None:
                            p.insert(0, copy.deepcopy(src_ppr))
                        r = etree.SubElement(p, f"{W}r")
                        if header_rpr is not None:
                            r.insert(0, copy.deepcopy(header_rpr))
                        t = etree.SubElement(r, f"{W}t")
                        t.text = "Rating"
                else:
                    set_grid_span(cell, span + 1)
                break
            pos += span

    return True


# ---------------------------------------------------------------------------
# Header cleanup: consolidate runs, remove junk chars, center-align
# ---------------------------------------------------------------------------

# Tahoma 10pt ≈ 100 DXA per character.
# Column width = (len(header) + 2) * DXA_PER_CHAR  (1 char padding each side)
DXA_PER_CHAR = 100
FIXED_WIDTH_HEADERS = {"result", "rating", "country"}

HEADER_NAMES = {
    "evaluation", "citation / method", "citation/method",
    "criteria", "country", "result", "rating",
}


def _is_data_table(tbl_el):
    """Return True if the table looks like a data table (has known header names)."""
    rows = tbl_el.findall(f"{W}tr")
    if not rows:
        return False
    texts = set()
    for cell in rows[0].findall(f"{W}tc"):
        texts.add(cell_text(cell).lower().strip())
    return bool(texts & HEADER_NAMES)


def _clean_text(raw: str) -> str:
    """Normalise whitespace: collapse runs of spaces/newlines into single space, strip."""
    import re
    return re.sub(r"\s+", " ", raw).strip()


def clean_header_cells(doc):
    """For every data-table header row: consolidate runs into clean text, center-align."""
    count = 0
    for table in doc.tables:
        tbl_el = table._tbl
        if not _is_data_table(tbl_el):
            continue
        rows = tbl_el.findall(f"{W}tr")
        if not rows:
            continue

        row0 = rows[0]
        for cell in row0.findall(f"{W}tc"):
            raw = cell_text(cell)
            if not raw:
                continue
            clean = _clean_text(raw)

            # Grab the rPr from the first run as a formatting template
            base_rpr = _copy_run_format(cell)
            # Strip character-spacing overrides from template
            if base_rpr is not None:
                sp = base_rpr.find(f"{W}spacing")
                if sp is not None:
                    base_rpr.remove(sp)

            # Remove ALL existing paragraphs
            for p in list(cell.findall(f"{W}p")):
                cell.remove(p)

            # Build a single clean paragraph: center-aligned, one run
            p = etree.SubElement(cell, f"{W}p")
            ppr = etree.SubElement(p, f"{W}pPr")
            jc = etree.SubElement(ppr, f"{W}jc")
            jc.set(f"{W}val", "center")

            r = etree.SubElement(p, f"{W}r")
            if base_rpr is not None:
                r.insert(0, copy.deepcopy(base_rpr))
            t = etree.SubElement(r, f"{W}t")
            t.text = clean
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        count += 1
    return count


# ---------------------------------------------------------------------------
# Font / size
# ---------------------------------------------------------------------------

def change_all_fonts(doc, name="Tahoma", size_pt=10):
    half_pt = str(size_pt * 2)
    body = doc.element.body

    for run in body.iter(f"{W}r"):
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.SubElement(run, f"{W}rPr")
            run.insert(0, rpr)

        rf = rpr.find(f"{W}rFonts")
        if rf is None:
            rf = etree.SubElement(rpr, f"{W}rFonts")
        rf.set(f"{W}ascii", name)
        rf.set(f"{W}hAnsi", name)
        rf.set(f"{W}eastAsia", name)
        rf.set(f"{W}cs", name)

        for tag in (f"{W}sz", f"{W}szCs"):
            el = rpr.find(tag)
            if el is None:
                el = etree.SubElement(rpr, tag)
            el.set(f"{W}val", half_pt)

    # Also update document default run properties (styles.xml)
    styles_el = doc.styles.element
    doc_defaults = styles_el.find(f"{W}docDefaults")
    if doc_defaults is None:
        doc_defaults = etree.SubElement(styles_el, f"{W}docDefaults")
        styles_el.insert(0, doc_defaults)
    rpr_default = doc_defaults.find(f"{W}rPrDefault")
    if rpr_default is None:
        rpr_default = etree.SubElement(doc_defaults, f"{W}rPrDefault")
    rpr = rpr_default.find(f"{W}rPr")
    if rpr is None:
        rpr = etree.SubElement(rpr_default, f"{W}rPr")

    rf = rpr.find(f"{W}rFonts")
    if rf is None:
        rf = etree.SubElement(rpr, f"{W}rFonts")
    rf.set(f"{W}ascii", name)
    rf.set(f"{W}hAnsi", name)
    rf.set(f"{W}eastAsia", name)
    rf.set(f"{W}cs", name)

    for tag in (f"{W}sz", f"{W}szCs"):
        el = rpr.find(tag)
        if el is None:
            el = etree.SubElement(rpr, tag)
        el.set(f"{W}val", str(size_pt * 2))


# ---------------------------------------------------------------------------
# Orientation & table auto-fit
# ---------------------------------------------------------------------------

def set_portrait(doc):
    for section in doc.sections:
        if section.orientation == WD_ORIENT.LANDSCAPE:
            w, h = section.page_width, section.page_height
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = min(w, h)
            section.page_height = max(w, h)


def _compute_uniform_col_widths(doc):
    """Compute a single set of column widths for all data tables.

    Fixed-width columns (FIXED_WIDTH_HEADERS) get their defined widths.
    Remaining columns share the leftover space proportionally, using the
    majority pattern from the source tables as a base.
    """
    W_ = W
    # Collect gridCol widths from every data table
    all_widths = []
    for table in doc.tables:
        tbl_el = table._tbl
        rows = tbl_el.findall(f"{W_}tr")
        if not rows:
            continue
        headers = [cell_text(c).lower() for c in rows[0].findall(f"{W_}tc")]
        if "evaluation" not in headers:
            continue
        tbl_grid = tbl_el.find(f"{W_}tblGrid")
        if tbl_grid is None:
            continue
        gcols = [int(gc.get(f"{W_}w", "0")) for gc in tbl_grid.findall(f"{W_}gridCol")]
        all_widths.append(tuple(gcols))

    if not all_widths:
        return None

    # Use most-common pattern as base
    from collections import Counter
    base = list(Counter(all_widths).most_common(1)[0][0])
    num_cols = len(base)

    # Determine which grid indices are fixed-width (by checking first table header)
    sample_tbl = None
    for table in doc.tables:
        tbl_el = table._tbl
        rows = tbl_el.findall(f"{W_}tr")
        if not rows:
            continue
        headers = [cell_text(c).lower() for c in rows[0].findall(f"{W_}tc")]
        if "evaluation" in headers:
            sample_tbl = tbl_el
            break

    fixed = {}  # grid_col_index → dxa
    if sample_tbl is not None:
        pos = 0
        for hcell in sample_tbl.findall(f"{W_}tr")[0].findall(f"{W_}tc"):
            span = get_grid_span(hcell)
            txt = cell_text(hcell).lower()
            if txt in FIXED_WIDTH_HEADERS:
                w = (len(txt) + 2) * DXA_PER_CHAR
                for gi in range(pos, pos + span):
                    fixed[gi] = w
            pos += span

    # Apply fixed widths into the base; redistribute remaining space
    fixed_total = sum(fixed.values())
    flex_indices = [i for i in range(num_cols) if i not in fixed]
    flex_original_total = sum(base[i] for i in flex_indices)

    # Desired total = original total (keeps proportions the same, Word scales to 100%)
    original_total = sum(base)
    flex_target = original_total - fixed_total

    uniform = list(base)
    for gi, w in fixed.items():
        if gi < num_cols:
            uniform[gi] = w
    if flex_original_total > 0 and flex_target > 0:
        for i in flex_indices:
            uniform[i] = round(base[i] / flex_original_total * flex_target)

    return uniform


def autofit_tables_to_window(doc):
    """Set every table to 100 % page width with uniform column widths for data tables."""
    uniform = _compute_uniform_col_widths(doc)

    for tbl in doc.element.body.iter(f"{W}tbl"):
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
            tbl.insert(0, tbl_pr)

        # Table width = 100 %
        tw = tbl_pr.find(f"{W}tblW")
        if tw is None:
            tw = etree.SubElement(tbl_pr, f"{W}tblW")
        tw.set(f"{W}w", "5000")
        tw.set(f"{W}type", "pct")

        # Remove fixed layout
        layout = tbl_pr.find(f"{W}tblLayout")
        if layout is not None:
            tbl_pr.remove(layout)

        # Check if this is a data table
        rows = tbl.findall(f"{W}tr")
        is_data = False
        if rows:
            headers = [cell_text(c).lower() for c in rows[0].findall(f"{W}tc")]
            is_data = "evaluation" in headers

        if is_data and uniform is not None:
            # Apply uniform gridCol widths
            tbl_grid = tbl.find(f"{W}tblGrid")
            if tbl_grid is not None:
                gcols = tbl_grid.findall(f"{W}gridCol")
                for gi, gc in enumerate(gcols):
                    if gi < len(uniform):
                        gc.set(f"{W}w", str(uniform[gi]))

            # Apply matching tcW on every cell
            for tr in rows:
                pos = 0
                for tc in tr.findall(f"{W}tc"):
                    span = get_grid_span(tc)
                    tc_pr = tc.find(f"{W}tcPr")
                    if tc_pr is not None:
                        tcw = tc_pr.find(f"{W}tcW")
                        if tcw is not None:
                            cell_w = sum(uniform[pos:pos + span]) if pos + span <= len(uniform) else 0
                            tcw.set(f"{W}type", "dxa")
                            tcw.set(f"{W}w", str(cell_w))
                    pos += span
        else:
            # Non-data tables: just set auto
            for tc in tbl.iter(f"{W}tc"):
                tc_pr = tc.find(f"{W}tcPr")
                if tc_pr is not None:
                    tcw = tc_pr.find(f"{W}tcW")
                    if tcw is not None:
                        tcw.set(f"{W}type", "auto")
                        tcw.set(f"{W}w", "0")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def remove_headers(doc):
    """Remove all header content from every section."""
    for section in doc.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            if header and header.is_linked_to_previous is False or header._element is not None:
                for p in list(header.paragraphs):
                    p._element.getparent().remove(p._element)
                for tbl in list(header.tables):
                    tbl._element.getparent().remove(tbl._element)
                header.is_linked_to_previous = True


def process(input_path: str, output_path: str):
    doc = Document(input_path)

    # 0. Remove headers
    remove_headers(doc)
    print("  Removed all headers")

    # 1. Delete columns
    cols_to_delete = ["No. of Samples", "Banner"]
    for col_name in cols_to_delete:
        count = 0
        for table in doc.tables:
            idx = find_column_index(
                table._tbl, lambda t, cn=col_name: cn.lower() in t.lower()
            )
            if idx >= 0:
                delete_column(table._tbl, idx)
                count += 1
        print(f"  Deleted '{col_name}' from {count} table(s)")

    # 2. Split Result/Rating
    count = 0
    for table in doc.tables:
        if split_result_rating(table._tbl):
            count += 1
    print(f"  Split 'Result / Rating' in {count} table(s)")

    # 3. Clean header cells (consolidate text, remove junk, center-align)
    hdr_count = clean_header_cells(doc)
    print(f"  Cleaned & centered headers in {hdr_count} table(s)")

    # 4. Fonts
    change_all_fonts(doc, "Tahoma", 10)
    print("  Changed all fonts to Tahoma 10pt")

    # 5. Orientation + auto-fit
    set_portrait(doc)
    autofit_tables_to_window(doc)
    print("  Set portrait orientation & auto-fit tables to window")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    print(f"  Saved → {output_path}")


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "source/原檔.docx"
    dst = sys.argv[2] if len(sys.argv) > 2 else "target/原檔_adjusted.docx"
    process(src, dst)
