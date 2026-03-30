# Document OCR to Excel Tracker

A Streamlit application that extracts data from BL (Bill of Lading) folder PDFs and automatically fills an Excel module tracker for solar module import transactions.

## Problem

Filling Excel trackers manually by cross-referencing 9+ PDF documents (invoices, customs forms, transport docs) is tedious and error-prone. This tool automates the entire process.

## How It Works

```
Upload BL Folder (PDFs) → Text Extraction → LLM Field Extraction → Review → Download Filled Excel
```

1. **Upload** - User uploads all PDFs from a BL folder (e.g., BL HANS003058)
2. **Text Extraction** - PyMuPDF extracts text from digital PDFs instantly. Scanned/image PDFs are sent to Gemini Flash vision model for OCR.
3. **Field Extraction** - Claude Sonnet reads the extracted text alongside Excel column headers and maps data to the correct fields as structured JSON.
4. **Merge & Validate** - Data from all 9 documents is merged, with cross-reference validation (invoice numbers, BL numbers, quantities match across documents).
5. **Download** - Filled Excel is generated from the fixed template, preserving formatting.

## Supported Documents

| # | Document | Type | Extraction Method |
|---|----------|------|-------------------|
| 1.1 | Supplier Commercial Invoice & Packing List | Scanned | Vision LLM |
| 1.2 | Certificate of Origin | Scanned | Vision LLM |
| 1.4 | Bill of Lading | Scanned | Vision LLM |
| 2 | HSS Agreement & Invoice | Scanned | Vision LLM |
| 3.1 | Bill of Entry (Customs) | Scanned | Vision LLM |
| 3.2 | Duty Payment Challan | Digital | PyMuPDF |
| 4.1 | E-Way Bill | Digital | PyMuPDF |
| 4.2 | Stamped LR Copy | Digital | PyMuPDF |
| 4.3 | GRN Copy | Digital | PyMuPDF |

## Project Structure

```
app.py                          # Streamlit UI
config.py                       # OpenRouter, model, template settings
requirements.txt                # Python dependencies
.env                            # Local OpenRouter API key (not committed)
.streamlit/secrets.toml.example # Streamlit Cloud secrets template

ocr/
    extractor.py                # PyMuPDF + Vision LLM text extraction

parsers/
    llm_parser.py               # OpenRouter LLM field extraction & merging
    prompts.py                  # Dynamic prompt builder
    validator.py                # Cross-reference validation

excel/
    reader.py                   # Dynamic Excel schema reader
    writer.py                   # Excel writer (preserves formatting)

templates/
    tracker_template.xlsx       # Blank Excel template (headers only)
```

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure API key

Add your OpenRouter API key to the `.env` file:

```
OPENROUTER_API_KEY=sk-or-v1-your-key-here
```

Get a key from [openrouter.ai/keys](https://openrouter.ai/keys).

For Streamlit Cloud, add the same key in the app's **Secrets** settings instead of creating a `.env` file:

```toml
OPENROUTER_API_KEY="sk-or-v1-your-key-here"
```

### 3. Run the app

```bash
streamlit run app.py
```

## Deploy To Streamlit Cloud

1. Push this project to GitHub.
2. Open [share.streamlit.io](https://share.streamlit.io/) and sign in with GitHub.
3. Click **New app** and select:
   - Repository: `arushigupta596/Renew_OCR`
   - Branch: `main`
   - Main file path: `app.py`
4. In **Advanced settings > Secrets**, add:

```toml
OPENROUTER_API_KEY="sk-or-v1-your-key-here"
```

5. Deploy the app.

Notes:
- `.env` is intentionally ignored and should stay local only.
- `Data/` is intentionally ignored so sample or sensitive documents are not published.
- `templates/tracker_template.xlsx` is included because the app needs it at runtime.

## Usage

1. Open the app in your browser (default: http://localhost:8501)
2. Upload all PDFs from a BL folder
3. Click **Extract Text from PDFs** - digital PDFs extract instantly, scanned ones go through vision OCR
4. Click **Extract Fields via LLM** - the LLM maps document content to Excel columns
5. Review the extracted data and validation results
6. Click **Generate Filled Excel** and download

## Models Used

| Task | Model | Why |
|------|-------|-----|
| OCR (scanned PDFs) | `google/gemini-2.0-flash-001` | Cheap, fast vision model |
| Field extraction & merging | `anthropic/claude-sonnet-4` | Accurate structured data extraction |

Both models are accessed via OpenRouter API.

## Configuration

Settings can be changed in `config.py`:

- `DEFAULT_MODEL` - LLM for field extraction (default: `anthropic/claude-sonnet-4`)
- `VISION_MODEL` - Vision model for scanned PDF OCR (default: `google/gemini-2.0-flash-001`)
- `AVAILABLE_MODELS` - Models shown in the sidebar dropdown

The LLM model can also be changed from the sidebar in the app.

## Excel Template

The tracker template (`templates/tracker_template.xlsx`) has 76 columns across 6 sections:

1. **Commercial Invoice Details** (A-K) - Supplier invoice, packing list, COO
2. **HSS Invoice** (L-T) - High Seas Sale invoice from EPC to SPV
3. **HSS Agreement** (U-Z) - Agreement references and BL linkage
4. **BoE Details** (AA-AQ) - Bill of Entry, duties, customs clearance
5. **E-Way Bill Details** (AR-BE) - Transport e-way bill, addresses, GST
6. **Transportation/Delivery** (BG-BV) - LR, vehicle, GRN, gate entry
7. **CIL Claim Estimate** (BX) - Final claim amount

## Dependencies

- **PyMuPDF** - Fast PDF text extraction
- **OpenAI SDK** - OpenRouter API calls (compatible)
- **openpyxl** - Excel read/write with formatting preservation
- **Streamlit** - Web UI
- **pandas** - Data display
- **python-dotenv** - Environment variable loading
