
import streamlit as st
import pandas as pd
import numpy as np
import re, io, zipfile, os, tempfile
from pathlib import Path
import fitz
from openpyxl import load_workbook
from copy import copy

st.set_page_config(page_title="Tata 1mg Precision Validation Tool", page_icon="🧪", layout="wide")

st.title("🧪 Tata 1mg Precision Validation Automation")
st.caption("INTRA / INTER precision population from Tata 1mg assay PDFs into your Excel template")

with st.sidebar:
    st.header("Study Configuration")
    level_mode = st.radio("QC Level Configuration", ["Auto Detect", "2 Levels", "3 Levels"])
    st.info("INTRA: 20 results per level. INTER Run 1 uses the first 5 INTRA results; subsequent INTER PDFs supply 5 results per run.")

st.header("1. Upload files")
template_file = st.file_uploader("Excel precision template (.xlsx)", type=["xlsx"])
intra_files = st.file_uploader("INTRA PDF(s)", type=["pdf"], accept_multiple_files=True)
inter_files = st.file_uploader("INTER PDF(s)", type=["pdf"], accept_multiple_files=True)

def pdf_text(uploaded):
    data = uploaded.getvalue()
    doc = fitz.open(stream=data, filetype="pdf")
    return "\n".join(page.get_text("text") for page in doc)

def normalize_analyte(x):
    x = x.strip()
    return re.sub(r'[^A-Za-z0-9]+', '_', x).strip('_').upper()

def parse_pdf(text):
    # Tata 1mg report records: PR INTRA/INTER <sample> <result> <unit> <RLU>
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    records=[]
    current_assay=None
    for i,line in enumerate(lines):
        m=re.search(r'Tests:\s*([^;]+)', line, re.I)
        if m:
            current_assay=normalize_analyte(m.group(1))
        elif current_assay and re.search(r'\bPR\s+(INTRA|INTER)\b', line, re.I):
            typ=re.search(r'\bPR\s+(INTRA|INTER)\b', line, re.I).group(1).upper()
            # Result is normally immediately after sample ID; use numeric value before unit.
            tail=line.split(typ,1)[-1].strip()
            nums=re.findall(r'(?<![A-Za-z])[-+]?\d+(?:\.\d+)?', tail)
            if not nums:
                continue
            # first number may be SID (e.g. 301); result generally appears after SID.
            # Tata reports use: PR INTER 301 2.5 g/dL ...
            if len(nums)>=2:
                result=float(nums[1])
                sid=nums[0]
            else:
                result=float(nums[0]); sid=""
            level_match=re.search(r'\b(L\d+)(?:\.\d+)?\b', tail, re.I)
            # For INTER reports, level is encoded by SID 301/302/303; for INTRA often L101/L201/L301 or L1.01...
            if level_match:
                level="L"+str(int(re.search(r'\d+', level_match.group(1)).group()))
            elif sid:
                d=re.search(r'(\d{3})$',sid)
                level="L"+d.group(1)[0] if d else None
            else:
                level=None
            records.append({"analyte":current_assay,"type":typ,"level":level,"result":result,"sid":sid})
    return pd.DataFrame(records)

def improve_levels(df):
    if df.empty: return df
    def fix(row):
        if row["level"] in ["L1","L2","L3"]: return row["level"]
        sid=str(row.get("sid",""))
        m=re.search(r'(\d{3})',sid)
        if m and m.group(1)[0] in "123": return "L"+m.group(1)[0]
        return row["level"]
    df["level"]=df.apply(fix,axis=1)
    return df

def find_target_cells(ws, level, section_hint):
    # Search text labels to find rows containing the requested level and section.
    hits=[]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str):
                s=c.value.upper().replace(" ","")
                if level.replace(" ","") in s and section_hint in s:
                    hits.append(c.coordinate)
    return hits

def write_results(ws, values, start_row, col):
    for idx,v in enumerate(values[:20]):
        ws.cell(start_row+idx, col).value = v

def populate_workbook(template_bytes, analyte_data, analyte):
    wb=load_workbook(io.BytesIO(template_bytes))
    # Work primarily on first sheet; preserve existing template sheets/styles/formulas.
    ws=wb[wb.sheetnames[0]]

    # Heuristic population: identify cells containing level labels, then populate numeric cells below/adjacent.
    # For custom templates, a review table is also included.
    # We create a dedicated machine-readable sheet without altering original layout.
    if "AUTO_DATA" in wb.sheetnames:
        del wb["AUTO_DATA"]
    data_ws=wb.create_sheet("AUTO_DATA")
    headers=["Analyte","Type","Level","Run","Replicate","Result"]
    data_ws.append(headers)

    for level in ["L1","L2","L3"]:
        intra=analyte_data[(analyte_data.type=="INTRA") & (analyte_data.level==level)]["result"].tolist()
        inter=analyte_data[(analyte_data.type=="INTER") & (analyte_data.level==level)]["result"].tolist()

        for r,v in enumerate(intra[:20],1):
            data_ws.append([analyte,"INTRA",level,"Within Run",r,v])

        # Run 1 = first 5 INTRA values
        for r,v in enumerate(intra[:5],1):
            data_ws.append([analyte,"INTER",level,"Run 1",r,v])

        # INTER PDFs are already combined here; if more than 5 values exist, group by 5.
        for j in range(0,len(inter),5):
            run=j//5+2
            for r,v in enumerate(inter[j:j+5],1):
                data_ws.append([analyte,"INTER",level,f"Run {run}",r,v])

    # Attempt to populate common template layout by finding level labels and writing down a nearby result column.
    # This is deliberately conservative; AUTO_DATA guarantees no loss of extracted data.
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str):
                txt=c.value.upper().replace(" ","")
                m=re.fullmatch(r'(?:LEVEL)?([123])',txt)
                if m:
                    level="L"+m.group(1)
                    vals=analyte_data[(analyte_data.type=="INTRA")&(analyte_data.level==level)]["result"].tolist()[:20]
                    if vals:
                        # Find first empty cell immediately below in same column.
                        rr=c.row+1
                        for v in vals:
                            while rr <= ws.max_row and ws.cell(rr,c.column).value not in (None,""):
                                rr+=1
                            if rr>ws.max_row: break
                            ws.cell(rr,c.column).value=v
                            rr+=1

    out=io.BytesIO()
    wb.save(out); out.seek(0)
    return out.getvalue()

if st.button("🔎 Analyze PDFs", type="primary", use_container_width=True):
    if not intra_files:
        st.error("Please upload at least one INTRA PDF.")
    elif not template_file:
        st.error("Please upload the Excel template.")
    else:
        all_intra=pd.concat([improve_levels(parse_pdf(pdf_text(f))) for f in intra_files], ignore_index=True)
        all_inter=pd.concat([improve_levels(parse_pdf(pdf_text(f))) for f in inter_files], ignore_index=True) if inter_files else pd.DataFrame(columns=all_intra.columns)
        all_data=pd.concat([all_intra,all_inter],ignore_index=True)

        if level_mode=="2 Levels":
            allowed=["L1","L2"]
        elif level_mode=="3 Levels":
            allowed=["L1","L2","L3"]
        else:
            found=sorted(set(x for x in all_data.level.dropna() if x in ["L1","L2","L3"]))
            allowed=found

        st.session_state["all_data"]=all_data
        st.session_state["allowed"]=allowed
        st.session_state["template"]=template_file.getvalue()

if "all_data" in st.session_state:
    df=st.session_state["all_data"]
    allowed=st.session_state["allowed"]
    analytes=sorted(df.analyte.dropna().unique())

    st.success(f"Detected {len(analytes)} analyte(s) and {len(allowed)} QC level(s): {', '.join(allowed)}")

    summary=[]
    for a in analytes:
        for l in allowed:
            i=len(df[(df.analyte==a)&(df.type=="INTRA")&(df.level==l)])
            e=len(df[(df.analyte==a)&(df.type=="INTER")&(df.level==l)])
            summary.append({"Analyte":a,"Level":l,"INTRA":i,"INTER PDF values":e,"INTRA status":"OK" if i>=20 else f"Need {20-i}","INTER status":"OK" if e>=5 or e==0 else "Review"})
    st.subheader("2. Review extracted data")
    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    st.subheader("3. Generate precision workbooks")
    if st.button("🚀 Generate All Analyte Excel Files", type="primary", use_container_width=True):
        template=st.session_state["template"]
        zip_buf=io.BytesIO()
        generated=[]
        with zipfile.ZipFile(zip_buf,"w",zipfile.ZIP_DEFLATED) as z:
            for a in analytes:
                ad=df[df.analyte==a].copy()
                data=ad[ad.level.isin(allowed)]
                xlsx=populate_workbook(template,data,a)
                fname=f"{a}_Precision.xlsx"
                z.writestr(fname,xlsx)
                generated.append(fname)
        zip_buf.seek(0)
        st.session_state["zip"]=zip_buf.getvalue()
        st.session_state["generated"]=generated
        st.success(f"Generated {len(generated)} analyte workbook(s).")

if "zip" in st.session_state:
    st.download_button("⬇️ DOWNLOAD ALL ANALYTES (ZIP)", st.session_state["zip"],
                       file_name="Tata1MG_Precision_Validation_Output.zip",
                       mime="application/zip", use_container_width=True)
    st.subheader("Generated files")
    for f in st.session_state["generated"]:
        st.write("✓", f)
