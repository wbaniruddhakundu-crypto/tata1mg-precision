import io
import re
import zipfile
from collections import defaultdict
from copy import copy

import fitz
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='Tata 1MG Precision Validation', page_icon='🧪', layout='wide')
st.title('🧪 Tata 1MG Precision Validation Automation')
st.caption('INTRA = 20 replicates/level | INTER Run 1 = INTRA first 5 | INTER Run 2+ = uploaded INTER PDFs in order')

NUM = r'[<>]?\s*-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?'
RECORD_RE = re.compile(rf'\bPR\s+(INTRA|INTER)\s+([A-Za-z0-9_.-]+)\s+({NUM})\s*([^\s]+)', re.I)

def normalize_analyte(name):
    name = re.sub(r'[^A-Za-z0-9]+', '_', str(name).strip())
    return re.sub(r'_+', '_', name).strip('_') or 'Analyte'

def level_from_sid(mode, sid):
    s = str(sid).upper().strip()
    if mode == 'INTRA':
        m = re.match(r'L([123])(?:[.]?01)?$', s) or re.search(r'L([123])', s)
        return int(m.group(1)) if m else None
    if re.fullmatch(r'[34]0[1-3]', s):
        return int(s[-1])
    m = re.fullmatch(r'R20([1-3])', s)
    if m:
        return int(m.group(1))
    m = re.search(r'0([1-3])$', s)
    return int(m.group(1)) if m else None

def extract_pdf_records(uploaded_file, expected_mode):
    doc = fitz.open(stream=uploaded_file.getvalue(), filetype='pdf')
    records = []
    for page_no, page in enumerate(doc, 1):
        text = page.get_text('text')
        if not text:
            continue
        m_assay = re.search(r'Tests:\s*([^;\n]+)', text, re.I)
        assay = m_assay.group(1).strip() if m_assay else None
        if not assay:
            lines = [x.strip() for x in text.splitlines() if x.strip()]
            for i, line in enumerate(lines):
                if line.lower().startswith('test status') and i + 1 < len(lines):
                    if lines[i + 1].lower() != 'assay report':
                        assay = lines[i + 1]
                        break
        for m in RECORD_RE.finditer(text):
            mode = m.group(1).upper()
            if mode != expected_mode:
                continue
            sid = m.group(2)
            raw = m.group(3).replace(' ', '')
            try:
                result = float(raw.replace('<', '').replace('>', ''))
            except ValueError:
                continue
            level = level_from_sid(mode, sid)
            if level is None:
                continue
            records.append({'analyte': assay or 'Unknown', 'mode': mode, 'sid': sid,
                            'level': level, 'result': result, 'unit': m.group(4), 'page': page_no})
    doc.close()
    return records

def group_records(records):
    grouped = defaultdict(lambda: defaultdict(list))
    units = {}
    for r in records:
        grouped[r['analyte']][r['level']].append(r['result'])
        units.setdefault(r['analyte'], r['unit'])
    return grouped, units

def find_layout(ws):
    # Default layout from the user's precision template: Inter B:D, Intra G:I, rows 7:26.
    layout = {'inter_cols':[2,3,4], 'intra_cols':[7,8,9], 'start_row':7, 'end_row':26,
              'n_row':28, 'mean_row':29, 'sd_row':30, 'cv_row':31,
              'level_header_row':6, 'replicate_col':6}
    for row in range(1, min(ws.max_row, 100) + 1):
        vals = [str(ws.cell(row, c).value or '').strip().lower() for c in range(1, min(ws.max_column, 15) + 1)]
        if 'replicate' in vals:
            layout['level_header_row'] = row
            layout['replicate_col'] = vals.index('replicate') + 1
        for c, v in enumerate(vals, 1):
            if v == 'n': layout['n_row'] = row
            elif v == 'mean': layout['mean_row'] = row
            elif v in {'sd','s.d.','s.d'}: layout['sd_row'] = row
            elif v == 'cv%': layout['cv_row'] = row
    return layout

def writable(ws, row, col):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                min_col, min_row, max_col, max_row = rng.bounds
                if row == min_row and col == min_col:
                    return ws.cell(min_row, min_col)
                raise ValueError(f'Merged cell {rng} blocks data cell {cell.coordinate}.')
    return cell

def put(ws, row, col, value):
    writable(ws, row, col).value = value

def copy_row_style(ws, source_row, target_row):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

def populate_workbook(template_bytes, analyte, unit, n_levels, intra_by_level, inter_runs_by_level):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    layout = find_layout(ws)

    # Number of INTER rows required = Run 1 (5) + one 5-replicate block per uploaded INTER PDF.
    max_inter_rows = max([len(intra_by_level.get(1, [])[:5])] +
                         [len(v[:5]) for vals in inter_runs_by_level.values() for v in vals]) if inter_runs_by_level else 5
    total_inter_rows = 5 + 5 * len(inter_runs_by_level.get(1, []))
    total_inter_rows = max(5, total_inter_rows)

    # Original template has 20 data rows. If more than 20 INTER observations are needed,
    # insert rows immediately before the statistics rows so every uploaded INTER PDF gets its own run.
    extra = max(0, total_inter_rows - (layout['end_row'] - layout['start_row'] + 1))
    if extra:
        stats_anchor = layout['n_row']
        ws.insert_rows(stats_anchor, amount=extra)
        for r in range(layout['end_row'] + 1, layout['end_row'] + extra + 1):
            copy_row_style(ws, layout['end_row'], r)
        layout['n_row'] += extra
        layout['mean_row'] += extra
        layout['sd_row'] += extra
        layout['cv_row'] += extra
        layout['end_row'] += extra

    # Clear all data cells in the expanded area. This removes dummy/template values from unused runs.
    for row in range(layout['start_row'], layout['end_row'] + 1):
        for col in layout['inter_cols'] + layout['intra_cols']:
            put(ws, row, col, None)
        put(ws, row, layout['replicate_col'], row - layout['start_row'] + 1)

    if ws['A3'].value:
        put(ws, 3, 1, f'Precision Studies of - {analyte}')
    if unit:
        put(ws, 7, 5, unit)

    # INTRA: exactly first 20 values per level.
    for level in range(1, 4):
        vals = list(intra_by_level.get(level, []))[:20]
        col = layout['intra_cols'][level - 1]
        for i, value in enumerate(vals):
            put(ws, layout['start_row'] + i, col, value)

    # INTER: Run 1 = INTRA first 5; each uploaded INTER PDF is the next run.
    for level in range(1, 4):
        col = layout['inter_cols'][level - 1]
        values = list(intra_by_level.get(level, []))[:5]
        for run_values in inter_runs_by_level.get(level, []):
            values.extend(list(run_values)[:5])
        for i, value in enumerate(values):
            put(ws, layout['start_row'] + i, col, value)

    # Rebuild N / Mean / SD / CV formulas after dynamic row expansion.
    s, e = layout['start_row'], layout['end_row']
    for col in layout['inter_cols'] + layout['intra_cols']:
        letter = get_column_letter(col)
        put(ws, layout['n_row'], col, f'=COUNT({letter}{s}:{letter}{e})')
        put(ws, layout['mean_row'], col, f'=IFERROR(AVERAGE({letter}{s}:{letter}{e}),"")')
        put(ws, layout['sd_row'], col, f'=IFERROR(STDEV({letter}{s}:{letter}{e}),"")')
        put(ws, layout['cv_row'], col, f'=IFERROR(({letter}{layout["sd_row"]}/{letter}{layout["mean_row"]})*100,"")')

    # For a 2-level study, remove level 3 values/formulas.
    if n_levels == 2:
        for col in (layout['inter_cols'][2], layout['intra_cols'][2]):
            for row in range(layout['start_row'], layout['cv_row'] + 1):
                put(ws, row, col, None)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = 'auto'
    except Exception:
        pass
    return wb

def wb_bytes(wb):
    b = io.BytesIO(); wb.save(b); return b.getvalue()

st.sidebar.header('Validation settings')
level_mode = st.sidebar.selectbox('QC level configuration', ['Auto Detect','2 Levels','3 Levels'])

st.info('Run mapping: Run 1 = first 5 INTRA. Run 2 = INTER PDF 1. Run 3 = INTER PDF 2. Run 4 = INTER PDF 3, and so on. No additional run is populated unless an INTER PDF is uploaded.')

template = st.file_uploader('1. Upload master Excel precision template', type=['xlsx','xlsm'])
intra_files = st.file_uploader('2. Upload INTRA PDF(s) — 20 replicates per level', type=['pdf'], accept_multiple_files=True)
inter_files = st.file_uploader('3. Upload INTER PDF(s) — upload in RUN order', type=['pdf'], accept_multiple_files=True)

if st.button('🔍 Process Validation Data', type='primary', use_container_width=True):
    if not template or not intra_files:
        st.error('Please upload the Excel template and at least one INTRA PDF.')
        st.stop()

    with st.spinner('Extracting PDF Result values...'):
        intra_records = []
        for f in intra_files:
            intra_records.extend(extract_pdf_records(f, 'INTRA'))
        intra_group, units = group_records(intra_records)

        inter_file_groups = []
        for f in inter_files or []:
            recs = extract_pdf_records(f, 'INTER')
            g, u = group_records(recs)
            inter_file_groups.append((f.name, g, u))

    if not intra_records:
        st.error('No INTRA Result records detected. Check the PDF format.')
        st.stop()

    detected_levels = sorted({lvl for a in intra_group.values() for lvl in a if lvl in (1,2,3)})
    n_levels = 2 if level_mode == '2 Levels' else 3 if level_mode == '3 Levels' else (3 if 3 in detected_levels else 2)

    st.subheader('Extraction summary')
    c1,c2,c3,c4 = st.columns(4)
    c1.metric('Analytes', len(intra_group)); c2.metric('QC levels', n_levels); c3.metric('INTRA records', len(intra_records)); c4.metric('INTER PDFs', len(inter_file_groups))

    summary = []
    for analyte in sorted(intra_group):
        for level in range(1, n_levels+1):
            counts = [len(g.get(analyte,{}).get(level,[])) for _,g,_ in inter_file_groups]
            summary.append({'Analyte':analyte,'Level':f'L{level}','INTRA':len(intra_group[analyte].get(level,[])),
                            'INTER PDF counts':', '.join(map(str,counts)) if counts else '-',
                            'Mapping':'Run 1 = INTRA first 5; Run 2+ = each INTER PDF'})
    st.dataframe(summary, use_container_width=True)

    outputs = {}; warnings=[]
    for analyte in sorted(intra_group):
        intra_by_level = {level:list(intra_group[analyte].get(level,[])) for level in range(1,n_levels+1)}
        for level in range(1,n_levels+1):
            if len(intra_by_level[level]) < 20:
                st.error(f'{analyte} L{level}: only {len(intra_by_level[level])}/20 INTRA results detected.')
                st.stop()

        inter_runs_by_level = defaultdict(list)
        # Every uploaded INTER PDF is one and only one new run.
        for filename, g, _ in inter_file_groups:
            for level in range(1,n_levels+1):
                vals = list(g.get(analyte,{}).get(level,[]))[:5]
                if len(vals) == 5:
                    inter_runs_by_level[level].append(vals)
                elif vals:
                    warnings.append(f'{analyte} L{level}: {filename} has {len(vals)}/5 INTER values; run left unpopulated.')
                else:
                    warnings.append(f'{analyte} L{level}: {filename} has no INTER values; run left unpopulated.')

        wb = populate_workbook(template.getvalue(), analyte, units.get(analyte,''), n_levels, intra_by_level, inter_runs_by_level)
        outputs[f'{normalize_analyte(analyte)}_Precision.xlsx'] = wb_bytes(wb)

    zip_b = io.BytesIO()
    with zipfile.ZipFile(zip_b,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in outputs.items(): z.writestr(name,data)
    zip_b.seek(0)

    st.success(f'Created {len(outputs)} analyte workbook(s).')
    if warnings:
        st.warning('Review these items:')
        for w in warnings: st.write('• '+w)
    st.download_button('⬇️ DOWNLOAD ALL ANALYTES — ZIP', data=zip_b.getvalue(), file_name='Tata1MG_Precision_Validation_Output.zip', mime='application/zip', use_container_width=True)
