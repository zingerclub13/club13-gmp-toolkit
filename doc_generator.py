"""
Document Generator for Club 13 GMP Toolkit

Generates .docx files from the actual Club 13 .dotx templates by
surgically filling values into the existing template structure
WITHOUT altering run boundaries, formatting, or layout.

Templates:
  - PK Specification Test Record Template.dotx
  - RM Specification Test Record Template.dotx
  - SOP Temp.dotx
"""
import io
import os
import re
import tempfile
import zipfile
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn

WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _val(row, key, default=""):
    """Safely get a value from a sqlite3.Row or dict."""
    try:
        v = row[key]
        return v if v is not None else default
    except (KeyError, IndexError):
        return default


# ── .dotx opener ────────────────────────────────────────────────

def open_dotx(path):
    """Open a .dotx template by patching its content type for python-docx."""
    with open(path, "rb") as f:
        data = io.BytesIO(f.read())

    patched = io.BytesIO()
    with zipfile.ZipFile(data, "r") as zin, zipfile.ZipFile(patched, "w") as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                content = content.replace(
                    b"wordprocessingml.template.main+xml",
                    b"wordprocessingml.document.main+xml",
                )
            zout.writestr(item, content)

    patched.seek(0)
    return Document(patched)


# ── Run-safe helpers ────────────────────────────────────────────

def _strip_fillin_fields(paragraph, field_values):
    """Remove FILLIN merge-field code runs from a paragraph and insert values.

    Word FILLIN fields consist of 3 consecutive runs:
      <w:r><w:fldChar fldCharType="begin"/></w:r>
      <w:r><w:instrText> FILLIN "Field Name" ...</w:instrText></w:r>
      <w:r><w:fldChar fldCharType="end"/></w:r>

    This function removes those runs and inserts a text run with the
    corresponding value from field_values dict after the label run.

    field_values maps the FILLIN prompt to the replacement text, e.g.:
      {"Component No.": "PKG-60W", "Rev. No.": "02", "Component Name": "Widget"}
    """
    p_elem = paragraph._element
    W = qn("w:r")
    runs_in_xml = list(p_elem.iterchildren(W))

    # Also look inside smartTags for runs
    for st in list(p_elem.iterchildren(qn("w:smartTag"))):
        runs_in_xml.extend(st.iterchildren(W))

    # Collect field code blocks: groups of (begin_run, instr_run, end_run)
    field_blocks = []
    i = 0
    all_children = list(p_elem)
    while i < len(all_children):
        child = all_children[i]
        if child.tag == qn("w:r"):
            fc = child.find(qn("w:fldChar"))
            if fc is not None and fc.get(qn("w:fldCharType")) == "begin":
                # Start of field — collect begin, instrText, end
                block_runs = [child]
                field_name = None
                j = i + 1
                while j < len(all_children):
                    sibling = all_children[j]
                    if sibling.tag == qn("w:r"):
                        block_runs.append(sibling)
                        # Check for instrText
                        instr = sibling.find(qn("w:instrText"))
                        if instr is not None and instr.text:
                            # Extract field name from: FILLIN "Field Name" \* MERGEFORMAT
                            import re as _re
                            m = _re.search(r'FILLIN\s+"([^"]+)"', instr.text)
                            if m:
                                field_name = m.group(1)
                        # Check for end
                        fc2 = sibling.find(qn("w:fldChar"))
                        if fc2 is not None and fc2.get(qn("w:fldCharType")) == "end":
                            field_blocks.append((block_runs, field_name))
                            i = j
                            break
                    j += 1
        i += 1

    # Remove field code runs and insert value runs
    for block_runs, field_name in field_blocks:
        # Find point to insert after (the last run before the field begin)
        insert_after = block_runs[0].getprevious()

        # Get the value
        value = ""
        if field_name and field_name in field_values:
            value = str(field_values[field_name])

        # Remove all field code runs
        for r in block_runs:
            r.getparent().remove(r)

        # Create a value run (inheriting formatting from the label run if available)
        new_r = p_elem.makeelement(qn("w:r"), {})
        new_t = new_r.makeelement(qn("w:t"), {})
        new_t.text = value
        new_t.set(qn("xml:space"), "preserve")
        new_r.append(new_t)

        # Insert after the label run
        if insert_after is not None:
            insert_after.addnext(new_r)
        else:
            p_elem.insert(0, new_r)


def _fill_after_label(paragraph, label, value):
    """Fill the first empty run after a label run, preserving all run boundaries.

    Template pattern:  [label_run] [empty_run] [empty_run] ...
    Result:            [label_run] [value_run] [empty_run] ...
    """
    runs = paragraph.runs
    label_found = False
    for run in runs:
        if label in run.text:
            label_found = True
            continue
        if label_found and run.text.strip() == "":
            run.text = str(value) if value else ""
            return True
    return False


def _replace_underscores(run, value):
    """Replace a contiguous underscore placeholder in a run with a value."""
    run.text = re.sub(r"_{3,}", str(value) if value else "", run.text, count=1)


def _set_row_cell_text(row_tr, col_idx, text):
    """Set text in a specific cell of a cloned row XML element."""
    cells = row_tr.findall(qn("w:tc"))
    if col_idx >= len(cells):
        return
    for p in cells[col_idx].findall(qn("w:p")):
        runs = p.findall(qn("w:r"))
        if runs:
            # Use existing first run — preserves its rPr
            for t in runs[0].findall(qn("w:t")):
                t.text = str(text) if text else ""
                t.set(qn("xml:space"), "preserve")
                return
            # Run exists but has no <w:t> — add one
            t_elem = runs[0].makeelement(qn("w:t"), {})
            t_elem.text = str(text) if text else ""
            t_elem.set(qn("xml:space"), "preserve")
            runs[0].append(t_elem)
            return
        # No runs at all — copy rPr from a sibling cell's run if available
        rPr_source = _find_sibling_rPr(row_tr, col_idx)
        r_elem = p.makeelement(qn("w:r"), {})
        if rPr_source is not None:
            r_elem.append(deepcopy(rPr_source))
        t_elem = r_elem.makeelement(qn("w:t"), {})
        t_elem.text = str(text) if text else ""
        t_elem.set(qn("xml:space"), "preserve")
        r_elem.append(t_elem)
        p.append(r_elem)
        return


def _find_sibling_rPr(row_tr, skip_col):
    """Find an rPr element from another cell in the same row for style copying."""
    for ci, tc in enumerate(row_tr.findall(qn("w:tc"))):
        if ci == skip_col:
            continue
        for p in tc.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    return rPr
    return None


# ══════════════════════════════════════════════════════════════════
# PK / RM SPECIFICATION TEST RECORD
# ══════════════════════════════════════════════════════════════════

def generate_pk_spec_record(template_path, spec, parameters, attachment_paths=None):
    """Generate a PK Specification Test Record with spec info filled in."""
    return _generate_spec_record(template_path, spec, parameters,
                                 attachment_paths=attachment_paths)


def generate_rm_spec_record(template_path, spec, parameters,
                            direct_params=None, coa_params=None,
                            attachment_paths=None):
    """Generate an RM Specification Test Record with spec info filled in.

    If direct_params/coa_params are provided, they are used for the two
    tables respectively.  Otherwise all parameters go into Table 0.
    """
    return _generate_spec_record(
        template_path, spec, parameters,
        direct_params=direct_params, coa_params=coa_params,
        attachment_paths=attachment_paths,
    )


def _generate_spec_record(template_path, spec, parameters,
                          direct_params=None, coa_params=None,
                          attachment_paths=None):
    """Shared logic for PK and RM spec record generation.

    Fills header fields (NO, Component No, Rev No, Component Name) and
    vendor line by surgically targeting the correct runs.
    Populates the Characteristic / Specifications table(s).
    Leaves Results / P / F columns blank for hand-filling.
    """
    doc = open_dotx(template_path)

    # ── Fill header paragraphs ──────────────────────────────────
    for section in doc.sections:
        hdr = section.header
        if not hdr:
            continue
        for p in hdr.paragraphs:
            text = p.text
            p_elem = p._element

            # Check for FILLIN merge fields in this paragraph
            has_fields = bool(p_elem.findall(f".//{{{W_NS}}}fldChar"))

            if "NO.:" in text and "____________" in text:
                # "LOT NO.: ____________" — replace underscores with spec number
                # Handle smartTag wrapping "LOT" by iterating all runs
                for run in p.runs:
                    if "____________" in run.text:
                        _replace_underscores(run, spec["spec_number"])

            elif "Component No.:" in text and "Rev. No.:" in text:
                if has_fields:
                    # Strip FILLIN field codes and replace with values
                    _strip_fillin_fields(p, {
                        "Component No.": _val(spec, "material_code"),
                        "Rev. No.": _val(spec, "revision", "00"),
                    })
                else:
                    _fill_after_label(p, "Component No.:", _val(spec, "material_code"))
                    _fill_after_label(p, "Rev. No.:", _val(spec, "revision", "00"))

            elif "Component Name:" in text:
                if has_fields:
                    _strip_fillin_fields(p, {
                        "Component Name": spec["material_name"],
                    })
                else:
                    _fill_after_label(p, "Component Name:", spec["material_name"])

    # ── Fill vendor in body (replace underscores in the correct runs) ──
    for p in doc.paragraphs:
        if "Vendor:" not in p.text:
            continue
        for run in p.runs:
            if "Vendor:" in run.text and "___" in run.text:
                _replace_underscores(run, _val(spec, "supplier"))
            # Don't touch "Vendor's Lot No." underscores — filled by hand

    # ── Fill test parameters table(s) ───────────────────────────
    tables = doc.tables
    if direct_params is not None or coa_params is not None:
        # RM with split parameters
        for table in tables:
            header_text = table.rows[0].cells[0].text.strip()
            if header_text == "Characteristic":
                _fill_spec_table(table, direct_params or [])
            elif "From C of A" in header_text:
                _fill_spec_table(table, coa_params or [])
    else:
        # PK (single table) or RM without split
        for table in tables:
            _fill_spec_table(table, parameters)

    # ── Append 3rd-party attachments as separate pages ────────────
    if attachment_paths:
        _append_attachments(doc, attachment_paths)

    # ── Save ────────────────────────────────────────────────────
    output = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    doc.save(output.name)
    return output.name


def _fill_spec_table(table, parameters):
    """Fill a Characteristic / Specifications table with parameter rows.

    Strategy:
      1. Deep-copy row 1 XML as the formatting template.
      2. Remove all non-header rows from the table.
      3. Append one cloned row per parameter in order.
    This preserves the exact row/cell formatting and keeps correct order.
    """
    if len(table.rows) < 2:
        return

    # Verify this is a spec table
    h0 = table.rows[0].cells[0].text.strip()
    if "Characteristic" not in h0 and "From C of A" not in h0:
        return

    # Save template row XML
    template_tr = deepcopy(table.rows[1]._tr)

    # Remove all non-header rows (iterate in reverse so indices stay valid)
    tbl_elem = table._tbl
    for ri in range(len(table.rows) - 1, 0, -1):
        tbl_elem.remove(table.rows[ri]._tr)

    if not parameters:
        # Keep one empty row so the table isn't just a header
        tbl_elem.append(deepcopy(template_tr))
        return

    # Append one row per parameter
    for param in parameters:
        new_tr = deepcopy(template_tr)
        _set_row_cell_text(new_tr, 0, param["parameter_name"])
        _set_row_cell_text(new_tr, 1, _val(param, "acceptance_criteria"))
        # Cols 2-3 (Results, Reference) left blank for hand-fill
        # Cols 4-5 (P, F) keep their template text for circling on paper
        tbl_elem.append(new_tr)


# ── Attachment appender ─────────────────────────────────────────

def _append_attachments(doc, attachment_paths):
    """Append 3rd-party attachments as separate pages in the document.

    attachment_paths: list of (original_name, file_path) tuples.

    - Images (.png, .jpg, .jpeg) are inserted full-width on a new page.
    - .docx files are merged paragraph-by-paragraph on a new page.
    - .pdf / other: a reference page is added noting the attached file.
    """
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_BREAK

    for original_name, file_path in attachment_paths:
        ext = os.path.splitext(file_path)[1].lower()

        # Add page break before each attachment
        bp = doc.add_paragraph()
        bp.runs[0].add_break(WD_BREAK.PAGE) if bp.runs else bp.add_run().add_break(WD_BREAK.PAGE)

        if ext in (".png", ".jpg", ".jpeg"):
            # Insert image — fit to page width
            heading = doc.add_paragraph()
            run = heading.add_run(f"Attachment: {original_name}")
            run.bold = True
            run.font.size = Pt(11)
            try:
                doc.add_picture(file_path, width=Inches(6.5))
            except Exception:
                doc.add_paragraph(f"[Image could not be embedded: {original_name}]")

        elif ext == ".docx":
            # Merge .docx content paragraph-by-paragraph
            heading = doc.add_paragraph()
            run = heading.add_run(f"Attachment: {original_name}")
            run.bold = True
            run.font.size = Pt(11)
            try:
                att_doc = Document(file_path)
                for p in att_doc.paragraphs:
                    new_p = doc.add_paragraph(p.text, style=p.style.name if p.style else None)
                    # Copy run-level formatting
                    if p.runs and new_p.runs:
                        for src_run, dst_run in zip(p.runs, new_p.runs):
                            if src_run.bold is not None:
                                dst_run.bold = src_run.bold
                            if src_run.italic is not None:
                                dst_run.italic = src_run.italic
                # Copy tables from attached doc
                for table in att_doc.tables:
                    _copy_table(doc, table)
            except Exception:
                doc.add_paragraph(f"[Document could not be embedded: {original_name}]")

        elif ext == ".pdf":
            heading = doc.add_paragraph()
            run = heading.add_run(f"Attachment: {original_name}")
            run.bold = True
            run.font.size = Pt(11)
            doc.add_paragraph(
                "This PDF attachment is included as a separate file. "
                "Please refer to the uploaded PDF document."
            )

        else:
            heading = doc.add_paragraph()
            run = heading.add_run(f"Attachment: {original_name}")
            run.bold = True
            run.font.size = Pt(11)
            doc.add_paragraph(f"[File type {ext} — see uploaded attachment]")


def _copy_table(doc, src_table):
    """Copy a table from one Document into another, preserving structure."""
    rows = len(src_table.rows)
    cols = len(src_table.columns)
    new_table = doc.add_table(rows=rows, cols=cols)
    for ri, row in enumerate(src_table.rows):
        for ci, cell in enumerate(row.cells):
            new_table.rows[ri].cells[ci].text = cell.text


# ══════════════════════════════════════════════════════════════════
# COMPONENT RECEIVING RECORD
# ══════════════════════════════════════════════════════════════════

def generate_receiving_record(template_path, spec, spec_type):
    """Generate a Component Receiving Record from the CLB003 template.

    The template has two identical pages (duplicate form).
    Fields filled from the spec:
      - COMPONENT CODE NO.  → spec.material_code
      - COMPONENT NAME      → spec.material_name
      - VENDOR              → spec.supplier
    All other fields (LOT NO., PO NO., DATE, etc.) are left blank for hand-fill.

    spec_type: 'pk' or 'rm' (used for filename context only).
    """
    doc = Document(template_path)

    field_map = {
        "COMPONENT CODE NO.:": _val(spec, "material_code"),
        "COMPONENT NAME:": _val(spec, "material_name"),
        "VENDOR:": _val(spec, "supplier"),
    }

    for p in doc.paragraphs:
        text = p.text
        for label, value in field_map.items():
            if label in text:
                _fill_underscore_field(p, label, value)

    output = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    doc.save(output.name)
    return output.name


def _fill_underscore_field(paragraph, label, value):
    """Replace the underscore placeholder after a label within a run's text.

    The template puts multiple fields in a single run, e.g.:
      'COMPONENT CODE NO.: __________  COMPONENT NAME: _____...'
    This function replaces only the underscore block immediately after the
    specified label, leaving other underscore blocks intact.
    """
    for run in paragraph.runs:
        if label in run.text:
            # Find the label position and the underscore block after it
            idx = run.text.index(label) + len(label)
            rest = run.text[idx:]
            m = re.search(r"_{3,}", rest)
            if m:
                before = run.text[:idx + m.start()]
                after = run.text[idx + m.end():]
                run.text = before + (str(value) if value else "") + after
            return


# ══════════════════════════════════════════════════════════════════
# SOP
# ══════════════════════════════════════════════════════════════════

def generate_sop(template_path, sop, revision_history=None):
    """Generate an SOP from the SOP Temp.dotx template.

    Fills:
      - Header text-box fields: SOP NO., REV NO.
      - Body section headings with content from the SOP record.
    Leaves signature table and footer untouched.
    """
    doc = open_dotx(template_path)

    # ── Fill header drawing text boxes ──────────────────────────
    for section in doc.sections:
        hdr = section.header
        if not hdr:
            continue
        # The SOP header uses wps:txbx text-box shapes (not tables)
        for txbx in hdr._element.iter(f"{{{WPS_NS}}}txbx"):
            for t_elem in txbx.iter(f"{{{W_NS}}}t"):
                if t_elem.text is None:
                    continue
                txt = t_elem.text.strip()
                if txt.startswith("SOP NO.:"):
                    t_elem.text = f"SOP NO.: {_val(sop, 'sop_number')}  "
                elif txt.startswith("REV NO.:"):
                    t_elem.text = f"REV NO.: {_val(sop, 'revision', '00')}"

    # ── Map section headings → content ──────────────────────────
    section_content = {
        "SUBJECT:": _val(sop, "title"),
        "OBJECTIVE:": _val(sop, "purpose"),
        "RESPONSIBILITIES:": _val(sop, "responsibilities"),
        "FREQUENCY:": _val(sop, "scope"),
        "NECESSARY EQUIPMENT:": _val(sop, "equipment_materials"),
        "PROCEDURE:": _val(sop, "procedure_text"),
        "DOCUMENTATION:": _val(sop, "references_text"),
    }

    # ── Fill sections ───────────────────────────────────────────
    for p in list(doc.paragraphs):
        text_upper = p.text.strip().upper()

        for heading, content in section_content.items():
            heading_upper = heading.upper()
            heading_bare = heading_upper.rstrip(": ")

            if text_upper in (heading_upper, heading_bare):
                if heading == "SUBJECT:":
                    # Append title on the same line, in the existing run
                    if p.runs:
                        # Keep "SUBJECT: " label, append title
                        p.runs[0].text = f"SUBJECT: {content}"
                elif content:
                    _insert_paragraphs_after(p, content)
                break

    # ── Save ────────────────────────────────────────────────────
    output = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    doc.save(output.name)
    return output.name


def _insert_paragraphs_after(ref_para, content):
    """Insert content paragraphs after a heading, using Normal style."""
    lines = content.strip().split("\n")

    for line in reversed(lines):
        new_p = ref_para._element.makeelement(qn("w:p"), {})

        # Set Normal style (not the heading's @Level 1)
        pPr = new_p.makeelement(qn("w:pPr"), {})
        pStyle = pPr.makeelement(qn("w:pStyle"), {})
        pStyle.set(qn("w:val"), "Normal")
        pPr.append(pStyle)
        new_p.append(pPr)

        # Add run with text
        r_elem = new_p.makeelement(qn("w:r"), {})
        t_elem = r_elem.makeelement(qn("w:t"), {})
        t_elem.text = line
        t_elem.set(qn("xml:space"), "preserve")
        r_elem.append(t_elem)
        new_p.append(r_elem)

        ref_para._element.addnext(new_p)
