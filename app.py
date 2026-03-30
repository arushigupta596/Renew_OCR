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
from parsers.validator import validate_cross_references
from excel.writer import write_to_tracker

st.set_page_config(page_title="OCR Document Tracker", layout="wide")
st.title("Document OCR → Excel Tracker")
st.caption("Upload a BL folder of PDF documents to extract data and generate the filled Excel tracker.")

# Config from .env
api_key = config.OPENROUTER_API_KEY
base_url = config.OPENROUTER_BASE_URL

# Load the fixed template
template_path = config.TEMPLATE_PATH
schema = read_tracker_schema(template_path)

# --- Sidebar: Settings ---
with st.sidebar:
    st.header("Settings")
    model = st.selectbox("LLM Model", config.AVAILABLE_MODELS, index=0)
    st.divider()
    overwrite = st.checkbox("Overwrite pre-filled cells", value=False)

# --- Step 1: Upload PDFs ---
st.header("Step 1: Upload BL Folder")

uploaded_pdfs = st.file_uploader(
    "Select all PDF files from the BL folder",
    type=["pdf"],
    accept_multiple_files=True,
    help="Select all PDFs inside your BL folder (e.g., BL HANS003058)",
)

# Save uploaded files to temp directory
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
    with st.expander("Uploaded PDFs"):
        for pdf_file in uploaded_pdfs:
            st.text(f"  {pdf_file.name}")

# --- Step 2: Text Extraction ---
st.header("Step 2: Extract Text from PDFs")

if "ocr_results" not in st.session_state:
    st.session_state.ocr_results = None

if st.button("Extract Text from PDFs", disabled=folder_path is None):
    progress = st.progress(0, text="Starting text extraction...")
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
        )
        st.session_state.ocr_results = results
        progress.progress(1.0, text="Text extraction complete!")

        errors = [f for f, t in results.items() if t.startswith("[OCR ERROR]")]
        if errors:
            st.warning(f"OCR errors on {len(errors)} files: {', '.join(errors)}")
        else:
            st.success(f"Successfully extracted text from {len(results)} PDFs")
    except Exception as e:
        st.error(f"OCR failed: {e}")

if st.session_state.ocr_results:
    st.subheader("OCR Results Preview")
    for filename, text in st.session_state.ocr_results.items():
        with st.expander(f"{filename} ({len(text)} chars)"):
            st.text_area("", text, height=300, key=f"ocr_{filename}", disabled=True)

# --- Step 3: LLM Field Extraction ---
st.header("Step 3: Extract Fields")

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
    )
    st.session_state.extracted_data = all_extracted

    progress.progress(0.95, text="Merging data from all documents...")
    merged = merge_all_extractions(
        all_extracted,
        column_headers,
        None,  # No existing data - using blank template
        api_key=api_key,
        model=model,
        base_url=base_url,
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
    st.subheader("Per-Document Extraction Results")
    for filename, data in st.session_state.extracted_data.items():
        with st.expander(f"{filename}"):
            if "_error" in data or "_parse_error" in data:
                st.error(data.get("_error", data.get("_parse_error")))
                if "_raw_response" in data:
                    st.code(data["_raw_response"])
            else:
                st.json(data)

if st.session_state.merged_data and "_parse_error" not in st.session_state.merged_data:
    st.subheader("Merged Data (Review Before Saving)")

    merged = st.session_state.merged_data
    shared = merged.get("shared_row", {})
    per_vehicle = merged.get("per_vehicle_rows", [])

    if shared:
        st.write("**Shared columns (same for all rows):**")
        shared_display = {k: v for k, v in shared.items() if not k.startswith("_")}
        df_shared = pd.DataFrame([shared_display]).T
        df_shared.columns = ["Value"]
        st.dataframe(df_shared, use_container_width=True)

    if per_vehicle:
        st.write(f"**Per-vehicle data ({len(per_vehicle)} rows):**")
        df_per = pd.DataFrame(per_vehicle)
        st.dataframe(df_per, use_container_width=True)

    # --- Step 4: Validation ---
    st.header("Step 4: Validation")
    if st.session_state.extracted_data:
        validations = validate_cross_references(st.session_state.extracted_data)
        if validations:
            for v in validations:
                icon = {"pass": "PASS", "warning": "WARNING", "fail": "FAIL"}.get(v["status"], "INFO")
                st.write(f"**[{icon}] {v['field']}**: {v['message']}")
        else:
            st.info("No cross-reference validations to show.")

    # --- Step 5: Download ---
    st.header("Step 5: Download Filled Excel")

    if st.button("Generate Filled Excel"):
        output_filename = "Module_Tracker_Filled.xlsx"
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

            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Filled Excel",
                    data=f.read(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.success("Excel generated successfully!")
        except Exception as e:
            st.error(f"Failed to write Excel: {e}")
