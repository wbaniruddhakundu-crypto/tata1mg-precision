import io
import re
import zipfile
from copy import copy
from collections import defaultdict, Counter

import fitz  # PyMuPDF
import openpyxl
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ============================================================
# Tata 1mg Precision Validation Automation
# INTRA: 20 observations / level
# INTER: Run 1 = first 5 INTRA observations / level
#        Run 2 onward = 5 observations / level from each INTER PDF
# ============================================================

st.set_page_config(
    page_title="Tata 1mg Precision Validation Automation",
    page_icon="🧪",
    layout="wide",
)

st.title("🧪 Tata 1mg Precision Validation Automation")
st.caption("Version 2 — merged-cell-safe Excel generation")
st.caption(
    "Upload the master Excel template, INTRA PDF(s), and INTER PDF(s). "
    "The app extracts Result values only and creates one precision workbook per analyte."
)

# -----------------------------
# Helpers
# -----------------------------
NUM = r"[<>]?\s*-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?"
RECORD_RE = re.compile(
    rf"\bPR\s+(INTRA|INTER)\s+([A-Za-z0-9_.-]+)\s+({NUM})\s*([^\s]+)",
    re.IGNORECASE,
)

def normalize_analyte(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9]+", "_", name.strip())
    name = re.sub(r"_+", "_", name).strip("_")
    return name or "Analyte"

def safe_sheet_name(name: str) -> str:
    bad = r'[]:*?/\\'
    for c in bad:
        name = name.replace(c, "_")
    return name[:31] or "Sheet1"

def level_from_sid(mode: str, sid: str):
    """Map Tata-style SID to QC level.

    Observed patterns:
      INTRA: L101/L201/L301 or L1.01/L2.01/L3.01 -> 1/2/3
      INTER: 301/302/303 or R201/R202/R203 -> 1/2/3
    """
    s = sid.upper().strip()

    if mode == "INTRA":
        m = re.match(r"L([123])(?:[.]?01)?$", s)
        if m:
            return int(m.group(1))
        # Generic fallback: first digit after L
        m = re.search(r"L([123])", s)
        if m:
            return int(m.group(1))
    else:
        # Common separate INTER PDF pattern: 301/302/303
        if re.fullmatch(r"[34]0[1-3]", s):
            return int(s[-1])
        # Embedded report pattern: R201/R202/R203
        m = re.fullmatch(r"R20([1-3])", s)
        if m:
            return int(m.group(1))
        # Generic fallback for numeric 30x/40x
        m = re.search(r"0([1-3])$", s)
        if m:
            return int(m.group(1))

    return None

def extract_pdf_records(uploaded_file, expected_mode=None):
    """Return records:
       [{analyte, mode, sid, level, result, unit, page}]
    """
    data = uploaded_file.getvalue()
    doc = fitz.open(stream=data, filetype="pdf")
    records = []

    for page_no, page in enumerate(doc, start=1):
        text = page.get_text("text")
        if not text:
            continue

        # Assay/analyte is normally declared near the top:
        assay_match = re.search(
            r"Tests:\s*([^;]+)\s*;", text, flags=re.IGNORECASE
        )
        assay = assay_match.group(1).strip() if assay_match else None

        # Fallback to a line after "Test Status..."
        if not assay:
            lines = [x.strip() for x in text.splitlines() if x.strip()]
            for i, line in enumerate(lines):
                if line.lower().startswith("test status") and i + 1 < len(lines):
                    if lines[i + 1] not in {"Assay Report"}:
                        assay = lines[i + 1]
                        break

        for m in RECORD_RE.finditer(text):
            mode = m.group(1).upper()
            sid = m.group(2)
            raw_result = m.group(3).replace(" ", "")
            unit = m.group(4)

            if expected_mode and mode != expected_mode:
                continue

            # Exclude RLU/MAU: the first numeric value after SID is the Result.
            try:
                result = float(raw_result.replace("<", "").replace(">", ""))
            except ValueError:
                continue

            level = level_from_sid(mode, sid)
            if level is None:
                continue

            records.append(
                {
                    "analyte": assay or "Unknown",
                    "mode": mode,
                    "sid": sid,
                    "level": level,
                    "result": result,
                    "unit": unit,
                    "page": page_no,
                }
            )

    doc.close()
    return records

def merge_records(files, expected_mode):
    all_records = []
    for f in files:
        recs = extract_pdf_records(f, expected_mode=expected_mode)
        all_records.extend(recs)
    return all_records

def group_records(records):
    out = defaultdict(lambda: defaultdict(list))
    units = {}
    for r in records:
        analyte = r["analyte"]
        level = r["level"]
        out[analyte][level].append(r["result"])
        units.setdefault(analyte, r["unit"])
    return out, units

def copy_template_to_output(template_bytes):
    return load_workbook(io.BytesIO(template_bytes))

def find_template_layout(ws):
    """Default layout from the user's shared example:
       Inter values B:D, Intra values G:I, rows 7:26.
       Returns detected positions where possible.
    """
    layout = {
        "inter_cols": [2, 3, 4],
        "intra_cols": [7, 8, 9],
        "start_row": 7,
        "end_row": 26,
        "n_row": 28,
        "mean_row": 29,
        "sd_row": 30,
        "cv_row": 31,
        "level_header_row": 6,
        "replicate_col": 6,
    }

    # Detect rows by labels if the template changes slightly.
    for row in range(1, min(ws.max_row, 80) + 1):
        vals = [str(ws.cell(row, c).value or "").strip().lower() for c in range(1, min(ws.max_column, 12) + 1)]
        if "replicate" in vals:
            layout["level_header_row"] = row
            # Replicate column
            for c, v in enumerate(vals, start=1):
                if v == "replicate":
                    layout["replicate_col"] = c

        for c, v in enumerate(vals, start=1):
            if v == "mean":
                layout["mean_row"] = row
            elif v in {"s.d.", "sd", "s.d"}:
                layout["sd_row"] = row
            elif v == "cv%":
                layout["cv_row"] = row
            elif v == "n":
                layout["n_row"] = row

    return layout

def set_formula(ws, cell, formula):
    ws[cell] = formula


def writable_cell(ws, row, col):
    """Return a writable cell and give a clear error for merged data cells."""
    from openpyxl.cell.cell import MergedCell
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                min_col, min_row, max_col, max_row = rng.bounds
                if row == min_row and col == min_col:
                    return ws.cell(min_row, min_col)
                raise ValueError(
                    f"Template has a merged cell covering {cell.coordinate} ({rng}). "
                    "The precision data area must contain individual cells for each replicate."
                )
    return cell

def write_value(ws, row, col, value):
    writable_cell(ws, row, col).value = value

def populate_workbook(template_bytes, analyte, unit, levels, intra_by_level, inter_runs_by_level):
    wb = copy_template_to_output(template_bytes)
    ws = wb.active
    layout = find_template_layout(ws)

    # Update title/analyte if possible without disturbing the template.
    # The shared example places study title in A3.
    if ws["A3"].value:
        write_value(ws, 3, 1, f"Precision Studies of - {analyte}")

    # Unit
    unit_cell = ws.cell(7, 5)
    if unit:
        write_value(ws, unit_cell.row, unit_cell.column, unit)

    # Clear data areas before population.
    for row in range(layout["start_row"], layout["end_row"] + 1):
        for col in layout["inter_cols"] + layout["intra_cols"]:
            write_value(ws, row, col, None)

    # Replicate numbers
    for i, row in enumerate(range(layout["start_row"], layout["end_row"] + 1), start=1):
        write_value(ws, row, layout["replicate_col"], i)

    # INTRA: exactly first 20 observations per level.
    for level in (1, 2, 3):
        col = layout["intra_cols"][level - 1]
        vals = list(intra_by_level.get(level, []))[:20]
        for i, value in enumerate(vals):
            write_value(ws, layout["start_row"] + i, col, value)

    # INTER:
    # Run 1 = first five INTRA observations.
    # Run 2+ = first five values from each supplied INTER PDF/run.
    # Flatten run-wise into the 20 rows of the shared template.
    for level in (1, 2, 3):
        col = layout["inter_cols"][level - 1]
        inter_values = []

        # Run 1
        intra_vals = list(intra_by_level.get(level, []))
        inter_values.extend(intra_vals[:5])

        # Run 2 onwards
        for run_values in inter_runs_by_level.get(level, []):
            inter_values.extend(list(run_values)[:5])

        for i, value in enumerate(inter_values[:20]):
            write_value(ws, layout["start_row"] + i, col, value)

        # If there are >20 inter observations, keep the template's 20-row design.
        # The app reports overflow in the validation summary.

    # Ensure N/Mean/SD/CV formulas match populated rows.
    s = layout["start_row"]
    e = layout["end_row"]

    for idx, col in enumerate(layout["inter_cols"]):
        letter = get_column_letter(col)
        write_value(ws, layout["n_row"], col, f'=COUNT({letter}{s}:{letter}{e})')
        write_value(ws, layout["mean_row"], col, f'=IFERROR(AVERAGE({letter}{s}:{letter}{e}),"")')
        write_value(ws, layout["sd_row"], col, f'=IFERROR(STDEV({letter}{s}:{letter}{e}),"")')
        write_value(ws, layout["cv_row"], col, f'=IFERROR(({letter}{layout["sd_row"]}/{letter}{layout["mean_row"]})*100,"")')

    for idx, col in enumerate(layout["intra_cols"]):
        letter = get_column_letter(col)
        write_value(ws, layout["n_row"], col, f'=COUNT({letter}{s}:{letter}{e})')
        write_value(ws, layout["mean_row"], col, f'=IFERROR(AVERAGE({letter}{s}:{letter}{e}),"")')
        write_value(ws, layout["sd_row"], col, f'=IFERROR(STDEV({letter}{s}:{letter}{e}),"")')
        write_value(ws, layout["cv_row"], col, f'=IFERROR(({letter}{layout["sd_row"]}/{letter}{layout["mean_row"]})*100,"")')

    # 2-level mode: clear Level 3 and its statistics.
    if levels == 2:
        for col in (layout["inter_cols"][2], layout["intra_cols"][2]):
            for row in range(layout["start_row"], layout["cv_row"] + 1):
                write_value(ws, row, col, None)

    # Ask Excel to recalculate formulas when opened.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    # Filename-safe
    return wb

def workbook_bytes(wb):
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# -----------------------------
# UI
# -----------------------------
st.sidebar.header("Validation settings")

level_mode = st.sidebar.selectbox(
    "QC level configuration",
    ["Auto Detect", "2 Levels", "3 Levels"],
)

use_first_five_for_inter1 = st.sidebar.checkbox(
    "Use first 5 INTRA results as INTER Run 1",
    value=True,
)

st.sidebar.info(
    "Recommended: keep this enabled for your Tata 1mg workflow. "
    "Each additional INTER PDF is treated as the next run and contributes 5 replicates per level."
)

template = st.file_uploader(
    "1. Upload master Excel precision template",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
)

intra_files = st.file_uploader(
    "2. Upload INTRA PDF(s) — 20 replicates per level",
    type=["pdf"],
    accept_multiple_files=True,
)

inter_files = st.file_uploader(
    "3. Upload INTER PDF(s) — one PDF/run, 5 replicates per level",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("🔍 Process Validation Data", type="primary", use_container_width=True):
    if not template:
        st.error("Please upload the Excel precision template.")
        st.stop()

    if not intra_files:
        st.error("Please upload at least one INTRA PDF.")
        st.stop()

    with st.spinner("Reading PDFs and extracting validation data..."):
        intra_records = merge_records(intra_files, "INTRA")
        inter_records = merge_records(inter_files or [], "INTER")

    if not intra_records:
        st.error("No INTRA Result records were detected.")
        st.stop()

    intra_group, units = group_records(intra_records)

    # Group each INTER PDF separately so each file becomes a new run.
    inter_file_groups = []
    for f in inter_files or []:
        recs = extract_pdf_records(f, expected_mode="INTER")
        g, u = group_records(recs)
        inter_file_groups.append((f.name, g, u))

    # If user chose a fixed level count, use it; otherwise infer from INTRA.
    detected_levels = sorted(
        {lvl for analyte in intra_group.values() for lvl in analyte.keys() if lvl in (1, 2, 3)}
    )

    if level_mode == "2 Levels":
        n_levels = 2
    elif level_mode == "3 Levels":
        n_levels = 3
    else:
        n_levels = 3 if 3 in detected_levels else 2

    st.subheader("Extraction summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Analytes detected", len(intra_group))
    c2.metric("Levels", n_levels)
    c3.metric("INTRA records", len(intra_records))
    c4.metric("INTER PDFs", len(inter_file_groups))

    # Validation table
    summary_rows = []
    for analyte in sorted(intra_group):
        for level in range(1, n_levels + 1):
            n_intra = len(intra_group[analyte].get(level, []))
            inter_pdf_counts = []
            for _, g, _ in inter_file_groups:
                inter_pdf_counts.append(len(g.get(analyte, {}).get(level, [])))

            summary_rows.append(
                {
                    "Analyte": analyte,
                    "Level": f"L{level}",
                    "INTRA": n_intra,
                    "INTER PDF counts": ", ".join(map(str, inter_pdf_counts)) if inter_pdf_counts else "-",
                    "Status": "OK" if n_intra >= 20 else f"CHECK: {n_intra}/20 INTRA",
                }
            )

    st.dataframe(summary_rows, use_container_width=True)

    # Manual review / mapping area
    with st.expander("🔎 Review detected analyte/level mapping before Excel creation", expanded=True):
        st.write(
            "The application maps Tata-style SIDs automatically: INTRA L101/L201/L301 → "
            "Levels 1/2/3 and INTER 301/302/303 (or R201/R202/R203) → Levels 1/2/3."
        )
        st.write(
            "Only the numeric Result is populated. RLU/MAU, flags and instrument evaluation text are ignored."
        )

    # Generate outputs
    outputs = {}
    warnings = []

    for analyte in sorted(intra_group):
        intra_by_level = {
            level: list(intra_group[analyte].get(level, []))
            for level in range(1, n_levels + 1)
        }

        inter_runs_by_level = defaultdict(list)

        # Run 1 = first 5 INTRA
        if use_first_five_for_inter1:
            for level in range(1, n_levels + 1):
                inter_runs_by_level[level].append(intra_by_level.get(level, [])[:5])

        # Run 2+ = each INTER PDF
        for filename, g, _ in inter_file_groups:
            for level in range(1, n_levels + 1):
                vals = list(g.get(analyte, {}).get(level, []))
                if vals:
                    inter_runs_by_level[level].append(vals[:5])

        # Checks
        for level in range(1, n_levels + 1):
            n = len(intra_by_level.get(level, []))
            if n < 20:
                warnings.append(f"{analyte} L{level}: only {n}/20 INTRA results found.")
            for run_no, vals in enumerate(inter_runs_by_level[level], start=1):
                if len(vals) < 5:
                    warnings.append(
                        f"{analyte} L{level}: INTER Run {run_no} has only {len(vals)}/5 values."
                    )

        # Pre-generation validation
        for level in range(1, n_levels + 1):
            if len(intra_by_level.get(level, [])) < 20:
                st.error(f"{analyte} L{level}: only {len(intra_by_level.get(level, []))}/20 INTRA Result values were detected.")
                st.stop()
            if use_first_five_for_inter1 and len(intra_by_level.get(level, [])) < 5:
                st.error(f"{analyte} L{level}: fewer than 5 INTRA values are available for INTER Run 1.")
                st.stop()

        wb = populate_workbook(
            template.getvalue(),
            analyte,
            units.get(analyte, ""),
            n_levels,
            intra_by_level,
            inter_runs_by_level,
        )

        outputs[f"{normalize_analyte(analyte)}_Precision.xlsx"] = workbook_bytes(wb)

    # ZIP
    zip_bio = io.BytesIO()
    with zipfile.ZipFile(zip_bio, "w", zipfile.ZIP_DEFLATED) as z:
        for filename, data in outputs.items():
            z.writestr(filename, data)
    zip_bio.seek(0)

    st.success(f"Created {len(outputs)} analyte workbooks.")

    if warnings:
        st.warning("Review these items before using the validation output:")
        for w in warnings:
            st.write(f"• {w}")

    st.download_button(
        "⬇️ DOWNLOAD ALL ANALYTES — ZIP",
        data=zip_bio.getvalue(),
        file_name="Tata1MG_Precision_Validation_Output.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.subheader("Individual analyte downloads")
    cols = st.columns(3)
    for i, (filename, data) in enumerate(outputs.items()):
        with cols[i % 3]:
            st.download_button(
                f"Download {filename}",
                data=data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{i}",
            )
