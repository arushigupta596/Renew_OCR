import os
import tempfile
import streamlit as st
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

import config
from ocr.extractor import extract_all_pdfs
from excel.reader import read_tracker_schema, read_existing_data, get_column_headers_list
from parsers.llm_parser import parse_all_documents, merge_all_extractions
from excel.writer import write_to_tracker

# ── Page Config ──
st.set_page_config(page_title="ReNew Document Tracker", layout="wide", initial_sidebar_state="expanded")

# ── CSS ──
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800&display=swap');
    @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL@24,400,1');

    /* ── Reset ── */
    html, body, [class*="css"], .stMarkdown, .stText, button, input, select, textarea {
        font-family: 'Montserrat', sans-serif !important;
    }
    #MainMenu, footer, header {visibility: hidden;}

    /* ── Page background ── */
    .stApp {
        background: linear-gradient(170deg, #f0f2f0 0%, #F6F6F6 30%, #f9faf9 100%);
    }

    /* ── Main container ── */
    .block-container {
        padding: 0 2rem 3rem 2rem !important;
        max-width: 1060px;
    }

    /* ── Header ── */
    .renew-header {
        background: linear-gradient(135deg, #2b2e30 0%, #3a3d40 50%, #313638 100%);
        padding: 1.4rem 2.2rem;
        border-radius: 0 0 18px 18px;
        margin: 0 -2rem 2rem -2rem;
        display: flex;
        align-items: center;
        gap: 1.1rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        position: relative;
        overflow: hidden;
    }
    .renew-header::after {
        content: '';
        position: absolute;
        top: 0; right: 0;
        width: 300px; height: 100%;
        background: linear-gradient(135deg, transparent 0%, rgba(114,191,68,0.06) 100%);
        pointer-events: none;
    }
    .renew-header h1 {
        color: #FFFFFF;
        font-weight: 800;
        font-size: 1.55rem;
        margin: 0;
        letter-spacing: -0.3px;
        border: none !important;
        padding: 0 !important;
    }
    .renew-header .subtitle {
        color: #a8d86e;
        font-size: 0.82rem;
        font-weight: 500;
        margin: 0.15rem 0 0 0;
        letter-spacing: 0.2px;
    }
    .renew-logo {
        background: linear-gradient(135deg, #72BF44 0%, #5fa639 100%);
        color: white;
        font-weight: 800;
        font-size: 0.85rem;
        padding: 0.5rem 0.95rem;
        border-radius: 10px;
        letter-spacing: 1.5px;
        white-space: nowrap;
        flex-shrink: 0;
        box-shadow: 0 2px 8px rgba(114,191,68,0.3);
        text-transform: uppercase;
    }

    /* ── Step badge ── */
    .step-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        background: linear-gradient(135deg, #72BF44 0%, #5fa639 100%);
        color: white;
        font-weight: 700;
        font-size: 0.65rem;
        padding: 0.25rem 0.7rem;
        border-radius: 20px;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        margin-bottom: 0.4rem;
        box-shadow: 0 2px 6px rgba(114,191,68,0.25);
    }
    .step-title {
        color: #2b2e30;
        font-weight: 700;
        font-size: 1.05rem;
        margin: 0.15rem 0 0.1rem 0;
        letter-spacing: -0.2px;
    }
    .step-desc {
        color: #808080;
        font-size: 0.8rem;
        margin: 0 0 0.4rem 0;
        line-height: 1.45;
    }

    /* ── Native Streamlit containers (border=True) ── */
    [data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border: 1px solid #e8eae8;
        border-radius: 16px !important;
        padding: 0.2rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.04);
        transition: box-shadow 0.2s ease;
    }
    [data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        box-shadow: 0 4px 20px rgba(0,0,0,0.07);
    }

    /* ── h2 / h3 overrides ── */
    .stApp h2 {
        color: #313638 !important;
        border: none !important;
        padding: 0 !important;
        margin: 0 !important;
        font-size: 1rem !important;
    }
    .stApp h3 {
        margin-top: 0.2rem !important;
        margin-bottom: 0.3rem !important;
        font-size: 0.95rem !important;
    }

    /* ── Buttons ── */
    .stButton > button {
        background: linear-gradient(135deg, #72BF44 0%, #5fa639 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.8rem !important;
        font-size: 0.82rem !important;
        letter-spacing: 0.3px !important;
        transition: all 0.25s ease !important;
        box-shadow: 0 2px 8px rgba(114,191,68,0.2) !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #8ec641 0%, #72BF44 100%) !important;
        color: white !important;
        box-shadow: 0 6px 20px rgba(114,191,68,0.35) !important;
        transform: translateY(-1px) !important;
    }
    .stButton > button:active {
        transform: translateY(0) !important;
    }
    .stButton > button:disabled {
        background: #e0e0e0 !important;
        color: #aaa !important;
        box-shadow: none !important;
        transform: none !important;
    }

    /* ── Download button ── */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #313638 0%, #2b2e30 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.8rem !important;
        font-size: 0.82rem !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.12) !important;
        transition: all 0.25s ease !important;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #72BF44 0%, #5fa639 100%) !important;
        box-shadow: 0 6px 20px rgba(114,191,68,0.35) !important;
        transform: translateY(-1px) !important;
    }

    /* ── Expanders ── */
    [data-testid="stExpander"] {
        background: #fafbfa;
        border-radius: 12px;
        border: 1px solid #eaecea;
        margin-top: 0.5rem;
    }
    [data-testid="stExpander"] summary {
        font-weight: 600;
        font-size: 0.85rem;
        color: #4a4a4a;
    }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] {
        border-radius: 12px;
    }
    [data-testid="stFileUploader"] section {
        border: 2px dashed #d4d8d4 !important;
        border-radius: 12px !important;
        background: #fafbfa !important;
        transition: border-color 0.2s ease;
        padding: 0.6rem !important;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #72BF44 !important;
        background: #f5f9f2 !important;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #313638 0%, #262829 100%);
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span {
        color: #e8e8e8 !important;
    }
    [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"] {
        background: transparent;
        border: none;
        box-shadow: none;
    }

    /* ── Radio buttons (case selector) ── */
    .stRadio > div[role="radiogroup"] {
        gap: 0.5rem;
    }
    .stRadio > div[role="radiogroup"] > label {
        background: #f5f7f5;
        border: 1.5px solid #e0e3e0;
        border-radius: 10px;
        padding: 0.6rem 1rem !important;
        transition: all 0.2s ease;
        font-size: 0.82rem;
    }
    .stRadio > div[role="radiogroup"] > label:hover {
        border-color: #72BF44;
        background: #f0f8ec;
    }
    .stRadio > div[role="radiogroup"] > label[data-checked="true"],
    .stRadio > div[role="radiogroup"] > label:has(input:checked) {
        border-color: #72BF44;
        background: linear-gradient(135deg, #f0f8ec 0%, #e8f5e0 100%);
        box-shadow: 0 2px 8px rgba(114,191,68,0.15);
    }

    /* ── Alerts ── */
    .stAlert {
        border-radius: 12px;
        font-size: 0.83rem;
    }
    [data-testid="stAlert"] {
        border-radius: 12px;
    }

    /* ── Progress bar ── */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #72BF44, #8ec641) !important;
        border-radius: 10px;
    }

    /* ── Dataframe ── */
    [data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid #eaecea;
    }

    /* ── Divider ── */
    hr {
        border-color: #eaecea !important;
        margin: 1rem 0 !important;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ──
st.markdown("""
<div class="renew-header">
    <div class="renew-logo">ReNew</div>
    <div>
        <h1>Document Tracker</h1>
        <p class="subtitle">Upload BL folder PDFs to extract data and generate the filled Excel tracker</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Config
api_key = config.OPENROUTER_API_KEY
base_url = config.OPENROUTER_BASE_URL

# ── Sidebar ──
with st.sidebar:
    st.markdown("### Settings")
    model = st.selectbox("LLM Model", config.AVAILABLE_MODELS, index=0)
    st.divider()
    overwrite = st.checkbox("Overwrite pre-filled cells", value=False)

# ═══════════════════════════════════════════════════════════════
# CASE SELECTION
# ═══════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown('<span class="step-badge">Configuration</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-title">Select Case Type</p>', unsafe_allow_html=True)
    st.markdown('<p class="step-desc">Choose the transaction type — this determines which Excel template and column schema to use.</p>', unsafe_allow_html=True)

    selected_case = st.radio(
        "Transaction type",
        options=list(config.CASE_TEMPLATES.keys()),
        index=0,
        horizontal=True,
        label_visibility="collapsed",
    )

# Reset state when case changes
if "selected_case" not in st.session_state:
    st.session_state.selected_case = selected_case
elif st.session_state.selected_case != selected_case:
    st.session_state.selected_case = selected_case
    st.session_state.ocr_results = None
    st.session_state.extracted_data = None
    st.session_state.merged_data = None
    st.rerun()

# Load template for selected case
template_path = config.CASE_TEMPLATES[selected_case]
schema = read_tracker_schema(template_path)

# ═══════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ═══════════════════════════════════════════════════════════════
st.markdown("")  # spacer

with st.container(border=True):
    st.markdown('<span class="step-badge">Step 1</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-title">Upload BL Folder</p>', unsafe_allow_html=True)
    st.markdown('<p class="step-desc">Select all PDF documents from a single BL folder (e.g. BL HANS003058).</p>', unsafe_allow_html=True)

    uploaded_pdfs = st.file_uploader(
        "Select PDF files",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

folder_path = None

if "temp_dir" not in st.session_state:
    st.session_state.temp_dir = None

if uploaded_pdfs:
    if st.session_state.temp_dir is None:
        st.session_state.temp_dir = tempfile.mkdtemp(prefix="ocr_bl_")

    pdf_dir = os.path.join(st.session_state.temp_dir, "pdfs")
    os.makedirs(pdf_dir, exist_ok=True)

    for pdf_file in uploaded_pdfs:
        pdf_path = os.path.join(pdf_dir, pdf_file.name)
        with open(pdf_path, "wb") as f:
            f.write(pdf_file.getbuffer())

    folder_path = pdf_dir
    st.success(f"Uploaded {len(uploaded_pdfs)} PDF files")
    with st.expander(f"View uploaded files ({len(uploaded_pdfs)})"):
        for pdf_file in uploaded_pdfs:
            st.text(f"  {pdf_file.name}")

# ═══════════════════════════════════════════════════════════════
# STEP 2 — Text Extraction
# ═══════════════════════════════════════════════════════════════
st.markdown("")  # spacer

with st.container(border=True):
    st.markdown('<span class="step-badge">Step 2</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-title">Extract Text from PDFs</p>', unsafe_allow_html=True)
    st.markdown('<p class="step-desc">Digital PDFs are extracted instantly via PyMuPDF. Scanned documents use Vision OCR (gpt-4o-mini).</p>', unsafe_allow_html=True)

    if "ocr_results" not in st.session_state:
        st.session_state.ocr_results = None

    if st.button("Extract Text", disabled=folder_path is None):
        progress = st.progress(0, text="Starting vision extraction...")
        status = st.empty()

        def ocr_progress(current, total, filename):
            progress.progress(current / total, text=f"Processing {filename} ({current}/{total})")
            status.text(f"Completed: {filename}")

        try:
            results = extract_all_pdfs(
                folder_path,
                api_key=api_key,
                vision_model=config.VISION_MODEL,
                base_url=base_url,
                progress_callback=ocr_progress,
                force_vision=False,
            )
            st.session_state.ocr_results = results
            progress.progress(1.0, text="Vision extraction complete!")

            errors = [f for f, t in results.items() if t.startswith("[OCR ERROR]")]
            if errors:
                st.warning(f"OCR errors on {len(errors)} files: {', '.join(errors)}")
            else:
                st.success(f"Successfully extracted text from {len(results)} PDFs")
        except Exception as e:
            st.error(f"OCR failed: {e}")

    if st.session_state.ocr_results:
        with st.expander(f"OCR Results Preview ({len(st.session_state.ocr_results)} documents)"):
            for filename, text in st.session_state.ocr_results.items():
                st.markdown(f"**{filename}** — {len(text):,} chars")
                st.text_area("", text, height=180, key=f"ocr_{filename}", disabled=True)
                st.divider()

# ═══════════════════════════════════════════════════════════════
# STEP 3 — LLM Field Extraction
# ═══════════════════════════════════════════════════════════════
st.markdown("")  # spacer

with st.container(border=True):
    st.markdown('<span class="step-badge">Step 3</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-title">Extract & Merge Fields</p>', unsafe_allow_html=True)
    st.markdown('<p class="step-desc">The LLM reads extracted text, maps content to Excel columns, and merges data across all documents.</p>', unsafe_allow_html=True)

    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = None
    if "merged_data" not in st.session_state:
        st.session_state.merged_data = None

    can_extract = st.session_state.ocr_results and api_key

    if st.button("Extract Fields via LLM", disabled=not can_extract):
        column_headers = get_column_headers_list(schema)

        progress = st.progress(0, text="Extracting fields...")

        def extract_progress(current, total, filename):
            progress.progress(current / total, text=f"Extracting from {filename} ({current}/{total})")

        all_extracted = parse_all_documents(
            st.session_state.ocr_results,
            column_headers,
            api_key=api_key,
            model=model,
            base_url=base_url,
            progress_callback=extract_progress,
            case=st.session_state.get("selected_case", ""),
        )
        st.session_state.extracted_data = all_extracted

        progress.progress(0.95, text="Merging data from all documents...")
        merged = merge_all_extractions(
            all_extracted,
            column_headers,
            None,
            api_key=api_key,
            model=model,
            base_url=base_url,
            case=st.session_state.get("selected_case", ""),
            ocr_results=st.session_state.ocr_results,
        )
        st.session_state.merged_data = merged
        progress.progress(1.0, text="Extraction complete!")

        if "_parse_error" in merged:
            st.error(f"Merge failed: {merged.get('_parse_error')}")
            with st.expander("Raw LLM response"):
                st.code(merged.get("_raw_response", ""))
        else:
            st.success("Fields extracted and merged successfully!")

    if st.session_state.extracted_data:
        with st.expander(f"Per-Document Extraction ({len(st.session_state.extracted_data)} documents)"):
            for filename, data in st.session_state.extracted_data.items():
                st.markdown(f"**{filename}**")
                if "_error" in data or "_parse_error" in data:
                    st.error(data.get("_error", data.get("_parse_error")))
                    if "_raw_response" in data:
                        st.code(data["_raw_response"])
                else:
                    st.json(data)
                st.divider()

# ═══════════════════════════════════════════════════════════════
# MERGED DATA REVIEW + STEP 4 DOWNLOAD
# ═══════════════════════════════════════════════════════════════
if st.session_state.get("merged_data") and "_parse_error" not in st.session_state.merged_data:
    st.markdown("")  # spacer

    merged = st.session_state.merged_data
    shared = merged.get("shared_row", {})
    per_vehicle = merged.get("per_vehicle_rows", [])

    with st.container(border=True):
        st.markdown('<span class="step-badge">Review</span>', unsafe_allow_html=True)
        st.markdown('<p class="step-title">Merged Data Preview</p>', unsafe_allow_html=True)
        st.markdown('<p class="step-desc">Review the extracted and merged data before generating the final Excel file.</p>', unsafe_allow_html=True)

        col1, col2 = st.columns([1, 1], gap="medium")

        with col1:
            if shared:
                st.markdown("**Shared Columns**")
                shared_display = {k: v for k, v in shared.items() if not k.startswith("_")}
                df_shared = pd.DataFrame([shared_display]).T
                df_shared.columns = ["Value"]
                st.dataframe(df_shared, use_container_width=True, height=320)

        with col2:
            if per_vehicle:
                st.markdown(f"**Per-Vehicle Data** ({len(per_vehicle)} rows)")
                df_per = pd.DataFrame(per_vehicle)
                st.dataframe(df_per, use_container_width=True, height=320)

    st.markdown("")  # spacer

    with st.container(border=True):
        st.markdown('<span class="step-badge">Step 4</span>', unsafe_allow_html=True)
        st.markdown('<p class="step-title">Download Filled Excel</p>', unsafe_allow_html=True)
        st.markdown('<p class="step-desc">Generate the completed tracker using the selected case template.</p>', unsafe_allow_html=True)

        col_btn, col_dl = st.columns([1, 1])

        with col_btn:
            if st.button("Generate Filled Excel"):
                case_short = selected_case.split("\u2013")[0].strip().replace(" ", "")
                output_filename = f"Module_Tracker_{case_short}_Filled.xlsx"
                output_path = os.path.join(st.session_state.temp_dir, output_filename)
                try:
                    write_to_tracker(
                        template_path=template_path,
                        output_path=output_path,
                        merged_data=merged,
                        schema=schema,
                        existing_data=[],
                        overwrite_existing=overwrite,
                    )
                    st.session_state.generated_excel = {
                        "path": output_path,
                        "filename": output_filename,
                    }
                    st.success("Excel generated!")
                except Exception as e:
                    st.error(f"Failed to write Excel: {e}")

        with col_dl:
            if st.session_state.get("generated_excel"):
                excel_info = st.session_state.generated_excel
                with open(excel_info["path"], "rb") as f:
                    st.download_button(
                        label="Download Excel",
                        data=f.read(),
                        file_name=excel_info["filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
