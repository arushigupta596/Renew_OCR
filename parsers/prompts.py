_CASE1_HINTS = """
## CASE 1 SPECIFIC GUIDANCE

### Document set in Case 1 PDF — extract from WHATEVER is present:
- Bills of Entry (BoE) — multiple (e.g., 6 BoEs filed against multiple Jinko invoices)
- E-Way Bills (EWB) — multiple (e.g., 3 EWBs)
- Domestic Tax Invoice (EPC → SPV) — 1 invoice
- Packing List, Certificate of Origin, Bill of Lading, LR (Lorry Receipt / Consignor Copy), GRN — extract if present; mark "No" / leave blank only if genuinely absent from the document

### Row structure:
- Create ONE ROW per Bill of Entry — each BoE is a separate row since each covers a different shipment
- Tax Invoice columns (AR–BG) repeat the SAME values across all rows (single domestic sale)
- E-Way Bill data goes into shared fields; link EWB to its corresponding BoE via Col AE

---

### SECTION 1 — Commercial Invoice (Columns C–O) — Source: Jinko Solar Commercial Invoice
- C  | Commercial Invoice No.: The invoice number declared in THIS BoE's Part IV (Invoice Details). One value per row — do NOT concatenate multiple invoice numbers. e.g. Z20240506319
- D  | Item Description as per Commercial Invoice: Product name, wattage, product code exactly as on the Jinko invoice — e.g. "SOLAR MODULE JKM580N-72HL4-BDV (580WP)"
- E  | Commercial Invoice Date: Date printed on the Jinko commercial invoice (DD-MMM-YYYY, e.g. 06-May-2024)
- F  | Bill To Party Name & Address: Buyer named on the Jinko invoice — typically ReNew Solar Energy entity with full address
- G  | Ship To Party Name & Address: Ship-to address on the Jinko invoice (same as Bill To in Case 1)
- H  | Quantity as per Commercial Invoice: Quantity of modules on the Jinko invoice that corresponds to THIS BoE (see Col AA for the matching invoice). Not the total across all BoEs. e.g. "18000 PCS (3600 x 580WP + 14400 x 585WP)"
- I  | Packing/Loading List Available: "Yes" if a packing list is in the document set, "No" if not present. If the document type is entirely absent from the PDF set, write "Not in PDF"
- J  | Certificate of Origin Available: "Yes" if a COO (Form AI / ASEAN-India) is present, "No" if not present. If entirely absent from PDF set, write "Not in PDF"
- K  | Invoice No on Packing List: Commercial invoice number cross-referenced on the Packing List. Write "Not in PDF" if no packing list exists
- L  | Invoice No on Certificate of Origin: Commercial invoice number on the COO. Write "Not in PDF" if no COO exists
- M  | Bill of Lading Available: "Yes" if an Ocean Bill of Lading is present, "No" if not present. If entirely absent from PDF set, write "Not in PDF"
- N  | BL Number: BL / HBL number from the shipping line. Write "Not in PDF" if no BL exists
- O  | Invoice No on Bill of Lading: Commercial invoice number on the BL. Write "Not in PDF" if no BL exists

### SECTION 2 — Bill of Entry (Columns P–AC) — Source: Indian Customs Bill of Entry (ICEGATE)
Each BoE gets its own row. Extract per-BoE fields from Part I, Part II and Part III of each BoE.
- P  | BOE Available: "Yes" — a BoE is present
- Q  | BoE Number: The BE number assigned by ICEGATE — e.g. 4539193. If multiple BoEs, each row has its own BoE number
- R  | Item Description as per BoE: Goods description from Part III of the BoE — e.g. "SOLAR MODULE JKM580N-72HL4-BDV (580WP), HSN 85414300"
- S  | BoE Date: The BE date printed on the BoE (DD-MMM-YYYY) — e.g. 16-Jul-2024
- T  | BoE Quantity: Commercial quantity declared in this BoE — include the wattage breakdown if shown. e.g. "18000 PCS (3600 x 580WP + 14400 x 585WP)"
- U  | BoE Assessable Amount: Total assessable value (ASS. VALUE) from Part III of the BoE — preserve the "INR" prefix and Indian number format exactly as printed. e.g. "INR 1,27,97,788.63"
- V  | BCD Amount: Basic Customs Duty from the duty summary. In Case 1, BCD is typically 0% — write "0 (BCD Rate = 0)" if BCD is nil/zero. Only write a numeric INR value if BCD is actually charged.
- W  | ADD Amount: Anti-Dumping Duty (ADD) from the duty summary of THIS specific BoE only. Extract the ADD figure from the duty calculation section of this BoE's pages — do not carry over ADD from a different BoE. Write "0 (ADD = 0)" if ADD is explicitly nil for this BoE.
- X  | BoE IGST Amount: IGST paid at import from the duty summary section of the BoE — preserve INR prefix and Indian number format. e.g. "INR 1,53,57,346.00"
- Y  | Duty Payment Challan Number: Challan number from the e-payment section of the BoE (Part I → F. Payment Details) — e.g. 2050099361. May be labeled "Challan No.", "CIN", or "e-Payment Transaction No."
- Z  | Duty Payment Challan Available: "Yes" if the duty payment challan/receipt is present
- AA | Commercial Invoice Number (linkage): The Jinko commercial invoice number as referenced inside the BoE (Part IV, Invoice Details) — links this BoE row to Column C
- AB | Seller Name & Address: Foreign exporter's name and address from BoE Part II, Section B — e.g. Jinko Solar (Vietnam) with full address
- AC | Buyer Name & Address: Indian importer's name and address from BoE Part II, Section B — e.g. ReNew Solar Energy (Jharkhand One) with full address

### SECTION 3 — E-Way Bill (Columns AD–AQ) — Source: E-Way Bill document
If multiple EWBs exist, list all EWB numbers in Col AD separated by "; ". Other EWB fields go in shared row.
- AD | E-Way Bill No.: 12-digit e-way bill number — e.g. 641756321324. List ALL if multiple, separated by "; "
- AE | BOE Number in E-Way Bill (linkage): The BoE number referenced in the EWB as the document basis — links to Col Q
- AF | E-way bill generated by: The entity that CREATED the EWB — whose GSTIN appears in the "Bill From" / header of the EWB. Typically a ReNew Solar Energy entity. Do NOT use the transporter name here. e.g. "RENEW SOLAR ENERGY (JHARKHAND ONE) PRIVATE LIMITED"
- AG | E-way bill generated by GSTIN: GSTIN of the EWB generator (Col AF). e.g. "08AAHCR7973H1ZS". This is NOT the transporter's GSTIN (which belongs in Col BP).
- AH | Item Description as per E-way Bill: Goods description from the EWB goods section — e.g. "SOLAR MODULE JKM580N-72HL4-BDV (580WP) & JKM585N-72HL4-BDV (585WP), HSN 85414300"
- AI | E-Way Bill Qty.: Quantity in the EWB goods table — often in KGS, e.g. 117400 KGS
- AJ | Type of E-way Bill: "Detailed" or "Normal" — imports must be Detailed
- AK | Bill From Name & Address: "Bill From" party in the EWB — e.g. Jinko Solar (Vietnam) Industries Co, Other Countries
- AL | Ship From Name & Address: "Dispatch From" address in the EWB — e.g. Adani Ports and SEZ, Mundra, Gujarat-370421
- AM | Bill To Name & Address: "Bill To" party in the EWB — e.g. ReNew Solar Energy (Jharkhand One), Rajasthan, GSTIN 08AAHCR7973H1ZS
- AN | Ship To Name & Address: "Ship To" destination in the EWB — e.g. SECI 8 Site Warehouse, Village Bogniyani, Tehsil Fatehgarh, Jaisalmer, Rajasthan-345027
- AO | Taxable Value / Value of Goods: Taxable amount from the EWB totals — e.g. 25586941.80
- AP | GST Amount: IGST amount from the EWB totals — e.g. 3070433.02
- AQ | Total Value (As per E-way Bill): Total invoice amount from the EWB totals — e.g. 28657374.82

### SECTION 4 — Tax Invoice (Columns AR–BG) — Source: ReNew domestic Tax Invoice (EPC → SPV)
Single Tax Invoice; same values repeated across all rows.
- AR | Tax Invoice No.: Invoice number on the domestic tax invoice issued by EPC to SPV — e.g. 5222360154
- AS | Item Description as per Invoice: Product description on the domestic invoice — e.g. "PV MODULE Bi-facial 580WP, HSN 85414300"
- AT | Bill From Party Name & Address: Seller on the tax invoice — ReNew Solar Energy (Jharkhand One) Pvt. Ltd. with GSTIN
- AU | Bill To Party Name & Address: Buyer on the tax invoice — e.g. ReNew Surya Vihaan Pvt Ltd with GSTIN
- AV | Ship To Party Name & Address (Consignee): Delivery address on the tax invoice — site/khasra address
- AW | Place of Supply: State and state code declared in the invoice — e.g. Rajasthan (State Code 08)
- AX | IRN No mentioned in Invoice: "Yes" if the Invoice Reference Number (IRN) from IRP is printed on the invoice
- AY | Tax Invoice Date: Date of the domestic tax invoice (DD-MMM-YYYY, e.g. 15-May-2024)
- AZ | Tax Invoice Quantity: Quantity billed on the domestic invoice — e.g. 82080 NOS
- BA | GST Rate: GST rate applied — e.g. 12% (CGST 6% + SGST 6%)
- BB | Taxable Value: Base taxable value before GST
- BC | GST Amount: Total GST charged (CGST + SGST or IGST)
- BD | TCS: Tax Collected at Source — 0.00 if not applicable
- BE | Invoice Value: Grand total of the Tax Invoice including GST
- BF | E-Invoice Copy Available: "Yes" if IRN + QR code are present on the invoice
- BG | Remarks: e.g. "EPC clears customs and sells domestically to SPV; 6 BOEs filed across multiple shipments"

### SECTION 5 — Transport / LR / GRN (Columns BI–BX) — Source: Lorry Receipt (Consignor Copy) and GRN
CRITICAL: A page or document labeled "Consignor Copy" IS the Lorry Receipt (LR). If ANY page in the document contains the text "Consignor Copy", that means an LR IS present — mark LR Available (Column BI) = "Yes" immediately.
LR Number (Column BN) = the Booking Number printed on that "Consignor Copy" page.
Extract all LR fields (BN through BQ) from the page labeled "Consignor Copy".
- BI | LR Available: "Yes" if ANY page is labeled "Consignor Copy" (that page IS the LR). "No" only if no such page exists anywhere in the document.
- BJ | Container number mentioned in LR Copy: Container ID from the Lorry Receipt
- BK | Invoice reference number (linkage): Commercial invoice number cross-referenced on the LR — links to Col C
- BL | E-way bill reference number (linkage): EWB number cross-referenced on the LR — links to Col AD
- BM | Vehicle No: Truck/vehicle registration number from the LR or EWB vehicle details
- BN | LR Number: The LR / consignment note number (Booking Number). PRIORITIZE from the page labeled "Consignor copy"
- BO | LR Date: Date of the Lorry Receipt (DD-MMM-YYYY). From the "Consignor copy" page
- BP | Transporter Name: Transport company name — e.g. Shamsons Tradelink LLP (may be visible in EWBs). This is NOT the EWB generator (Col AF).
- BQ | LR - Consignee Name & Address: Consignee name and delivery address on the LR
- BR | Gate Entry No.: Site gate entry number when the vehicle entered the project site
- BS | Gate Inward Date: Date of gate entry at the project site (DD-MMM-YYYY)
- BT | Gate Stamp Yes/No: "Yes" if a GATE IN STAMP is present on the gate entry document
- BU | GRN Available: "Yes" if a Goods Receipt Cum Inspection Report (GRN) is present, "No" if not
- BV | Invoice No on GRN: Commercial invoice number on the GRN — links to Col C
- BW | E-Way Bill No on GRN: EWB number on the GRN — links to Col AD
- BX | Vehicle No on GRN: Vehicle/truck number recorded on the GRN
"""

_CASE2_HINTS = """
## CASE 2 SPECIFIC GUIDANCE

### Document set in Case 2 PDF:
- Jinko commercial invoice + packing list + COO + BL are present
- HSS Invoice and HSS Agreement are present
- BoE is filed by the SPV/importer, not by ReNew EPC
- E-Way Bill, LR copies, GRN, and duty payment receipt are present

### Row structure:
- Create ONE ROW for the shipment/BoE in this example set
- Case 2 is not a Case 1-style multi-BoE tracker unless the PDF actually contains multiple BoEs

### Commercial Invoice (C-L):
- Use the foreign supplier invoice from Jinko, not the HSS invoice
- Bill To / Ship To on the Jinko invoice are ReNew Solar Energy (Jharkhand One) Pvt Ltd
- Packing List and COO are present, so I/J/K/L should be populated from those docs

### HSS Invoice + Agreement (M-AA):
- HSS invoice is ReNew EPC -> IB Vogt SPV
- Place of supply comes from the HSS invoice
- HSS agreement fields must come from the agreement, especially reference number, invoice references, BL number, and buyer details

### BL / BoE / EWB:
- BL is present and must be marked Yes
- BoE importer/buyer is IB VOGT SOLAR SEVEN PRIVATE LIMITED, not ReNew EPC
- E-Way Bill is generated by the SPV/importer IB VOGT SOLAR SEVEN PRIVATE LIMITED and GSTIN 08AAFCI4907A1ZX

### LR / GRN:
- LR is available if any consignor copy / lorry receipt is present
- GRN is available if Goods Receipt Cum Inspection Report is present
- Aggregate multiple LR / container / vehicle / gate-entry values with comma-separated strings when they belong to the same shipment row

### Preserve formatting:
- Dates: DD-MMM-YYYY
- Amounts in amount-specific fields should preserve INR prefix where shown in the source or where clearly monetary in the tracker context
- Do not mix up EPC and SPV roles across HSS / BoE / EWB sections
"""

_CASE3_HINTS = """
## CASE 3 SPECIFIC GUIDANCE

### Document set in Case 3 PDF:
- Domestic supplier tax invoices from ReNew Photovoltaics Pvt Ltd
- E-Way Bills for outward domestic supply
- FG Vehicle Loading Checklists / packing-loading records
- Lorry Receipts from multiple transporters
- GRNs / goods receipt records at site
- One consolidated EPC -> SPV domestic tax invoice

### Row structure:
- Create ONE ROW per domestic supplier invoice / truck load
- In the example set there are 9 supplier invoices, each typically linked to one EWB, one LR, one vehicle, and one GRN
- Columns AH-AV are the consolidated EPC -> SPV invoice block and repeat across all rows

### Critical mapping:
- Columns C-S come from the domestic supplier tax invoice + loading checklist
- Columns T-AG come from the E-Way Bill linked to that supplier invoice
- Columns AH-AW come from the single consolidated EPC -> SPV invoice
- Columns AY-BN come from the LR / transporter copy / GRN linked to that supplier invoice

### Domestic transaction rules:
- This is a fully domestic case. No foreign supplier, COO, BL, or BoE/customs fields apply here.
- Packing/Loading List Available should be "Yes" when the FG Vehicle Loading Checklist is present.
- E-Way Bill generator is the domestic supplier ReNew Photovoltaics Pvt Ltd.
- Keep the invoice-to-EWB-to-LR linkage consistent row by row.

### Preserve formatting:
- Case 3 trackers often store dates as DD.MM.YYYY strings. Preserve the date style visible in the source if the tracker values use dots.
- Keep transporter / vehicle / GRN values row-specific. Do not aggregate all 9 truck loads into one row.
"""


def build_extraction_prompt(
    ocr_text: str,
    column_headers: list[str],
    doc_filename: str,
    case: str = "",
) -> str:
    """
    Build a prompt that asks the LLM to extract data from OCR text
    and map it to Excel column headers.
    """
    headers_str = "\n".join(column_headers)
    if "Case 1" in case:
        case_hints = _CASE1_HINTS
    elif "Case 2" in case:
        case_hints = _CASE2_HINTS
    elif "Case 3" in case:
        case_hints = _CASE3_HINTS
    else:
        case_hints = ""

    return f"""You are a document data extraction expert. You are given OCR-extracted text from a PDF document and a list of Excel column headers. Your job is to extract values from the document that correspond to the column headers.

## Document filename: {doc_filename}

## Excel Column Headers (extract values for these):
{headers_str}

## Rules:
1. Return a JSON object where keys are the column identifiers (e.g., "Column B (2)") and values are the extracted data.
2. Only include columns where you can find a corresponding value in the document. Skip columns that don't relate to this document.
3. For dates, return in DD-MMM-YYYY format (e.g., 06-May-2024). Never use YYYY-MM-DD.
4. For monetary amounts in BoE duty/value columns (U, V, W, X): preserve the INR prefix and Indian number format exactly as printed (e.g., "INR 1,27,97,788.63"). For all other numeric fields, return the plain number without currency symbols.
5. For Yes/No fields, return "Yes" or "No".
6. For text fields with addresses, preserve line breaks as \\n.
7. If a field is clearly present but the value is not applicable, use "N/A".
8. Do NOT guess or fabricate values. Only extract what is clearly stated in the document.
{case_hints}
## OCR-Extracted Document Text:
{ocr_text}

## Response:
Return ONLY a valid JSON object. No markdown formatting, no explanation, just the JSON."""


def build_multi_row_extraction_prompt(
    ocr_text: str,
    column_headers: list[str],
    doc_filename: str,
    case: str = "",
) -> str:
    """
    Build a prompt for documents that may contain per-vehicle/per-item data
    (e.g., E-Way Bill vehicle list, LR copies, GRN with multiple entries).
    """
    headers_str = "\n".join(column_headers)
    if "Case 1" in case:
        case_hints = _CASE1_HINTS
    elif "Case 2" in case:
        case_hints = _CASE2_HINTS
    elif "Case 3" in case:
        case_hints = _CASE3_HINTS
    else:
        case_hints = ""

    return f"""You are a document data extraction expert. You are given OCR-extracted text from a PDF document and a list of Excel column headers.

This document may contain data for MULTIPLE rows (e.g., a list of vehicles, multiple lorry receipts, multiple gate entries).

## Document filename: {doc_filename}

## Excel Column Headers:
{headers_str}

## Rules:
1. Return a JSON object with two keys:
   - "shared": A JSON object of column values that are the SAME for all rows (e.g., e-way bill number, total quantities, addresses).
   - "per_row": A JSON array of objects, one per row/vehicle/entry. Each object contains column values unique to that row. Include a "vehicle_no" or identifier field if available.
2. Use column identifiers as keys (e.g., "Column B (2)").
3. For dates, return in DD-MMM-YYYY format (e.g., 06-May-2024). Never use YYYY-MM-DD.
4. For monetary amounts in BoE duty/value columns (U, V, W, X): preserve the INR prefix and Indian number format exactly as printed (e.g., "INR 1,27,97,788.63"). For all other numeric fields, return the plain number without currency symbols.
5. Do NOT guess or fabricate values.
{case_hints}
## OCR-Extracted Document Text:
{ocr_text}

## Response:
Return ONLY a valid JSON object with "shared" and "per_row" keys. No markdown, no explanation."""


_CASE1_MERGE_HINTS = """
## CASE 1 MERGE GUIDANCE — CRITICAL

### Row structure (non-negotiable):
- `per_vehicle_rows` MUST contain EXACTLY ONE ENTRY PER BILL OF ENTRY (BoE).
- Each BoE is identified by its unique BoE Number (Column Q). Typically 6 BoEs → 6 entries.
- The concept of "per vehicle" does NOT apply here — BoE Number is the unique row identifier.
- Do NOT collapse multiple BoEs into one row. Do NOT concatenate values across BoEs.

### What goes in each `per_vehicle_rows` entry (one dict per BoE):
- BoE fields: Columns P–AC (P, Q, R, S, T, U, V, W, X, Y, Z, AA, AB, AC)
- The EWB that references this BoE number in its document basis (AE linkage): Columns AD–AQ
- LR/transport data linked to this BoE/EWB: Columns BI–BX
- Column C (Invoice No.) for THIS row = the invoice number in Col AA of this BoE (one value only, not concatenated)

### What goes in `shared_row`:
- Commercial Invoice fields: Columns C–O (Jinko invoice data, same values across all BoE rows)
- Domestic Tax Invoice fields: Columns AR–BG (single domestic sale, repeated every row)

### EWB generator vs transporter:
- Column AF = entity that CREATED the EWB (typically a ReNew Solar Energy entity). Column AG = that entity's GSTIN.
- Column BP = transporter company name (e.g., Shamsons Tradelink LLP). Never put transporter name in AF or transporter GSTIN in AG.

### LR/GRN fields (Columns BI–BX):
- If ANY page is labeled "Consignor Copy", set BI = "Yes" and fill BJ–BQ from that page.
- Do NOT leave BI–BX empty if LR or GRN data was extracted from any document. Place these in the matching BoE row.

### Preserve string values:
- Dates must remain as DD-MMM-YYYY text strings (e.g., "06-May-2024"). Do not convert to numbers.
- INR-prefixed amounts must remain as strings (e.g., "INR 1,27,97,788.63"). Do not convert to numbers.
"""

_CASE2_MERGE_HINTS = """
## CASE 2 MERGE GUIDANCE

- Return a single shipment row unless the PDF clearly contains multiple distinct BoEs/shipments.
- Columns C-L must come from the Jinko supplier invoice / packing list / COO / BL set.
- Columns M-AA must come from the HSS invoice and HSS agreement only.
- Columns AB-AD come from the Bill of Lading.
- Columns AE-AR come from the BoE and duty payment receipt. The importer/buyer is the SPV.
- Columns AS-BF come from the E-Way Bill. The EWB generator is the SPV/importer, not ReNew EPC.
- Columns BH-BW come from LR copies and GRN. If multiple LRs exist for the same shipment, aggregate values using comma-separated strings in one row.
- Mark available documents as Yes when present in the PDF; do not mark COO/BL/LR/GRN as absent when they are present in OCR text.
- Keep party roles consistent:
  ReNew EPC = HSS seller / original foreign-invoice buyer
  IB Vogt SPV = HSS buyer, BoE importer, EWB generator, final consignee
"""

_CASE3_MERGE_HINTS = """
## CASE 3 MERGE GUIDANCE

- Return one row per domestic supplier invoice / truck load, not one row for the whole PDF.
- Columns C-S must come from the ReNew Photovoltaics invoice and its loading checklist.
- Columns T-AG must come from the E-Way Bill linked to that supplier invoice.
- Columns AH-AW come from the single consolidated EPC -> SPV invoice and repeat across all rows.
- Columns AY-BN must come from the LR / transporter copy / GRN linked to that same supplier invoice.
- Keep invoice number, EWB number, LR number, vehicle number, and GRN linkage aligned within each row.
- Do not mix transporter names across rows; each row should reflect the transporter for that truck.
- Do not collapse all LRs or all EWBs into one row for Case 3.
"""


def build_merge_prompt(
    all_extracted: dict[str, dict],
    column_headers: list[str],
    existing_row_data: list = None,
    case: str = "",
) -> str:
    """
    Build a prompt to merge extracted data from multiple documents into final row data.
    """
    headers_str = "\n".join(column_headers)

    docs_str = ""
    for filename, data in all_extracted.items():
        docs_str += f"\n### From {filename}:\n{_format_dict(data)}\n"

    existing_str = ""
    if existing_row_data:
        existing_str = f"""
## Existing Data in Excel (already filled - DO NOT overwrite unless the new value is clearly more complete):
{_format_existing(existing_row_data)}
"""

    if "Case 1" in case:
        case_merge_hints = _CASE1_MERGE_HINTS
    elif "Case 2" in case:
        case_merge_hints = _CASE2_MERGE_HINTS
    elif "Case 3" in case:
        case_merge_hints = _CASE3_MERGE_HINTS
    else:
        case_merge_hints = ""

    return f"""You are a data reconciliation expert. You have extracted data from multiple PDF documents that all belong to the same shipment/transaction. Merge them into a complete dataset for the Excel tracker.

## Excel Column Headers:
{headers_str}

## Extracted Data from Each Document:
{docs_str}
{existing_str}
{case_merge_hints}

## Rules:
1. Merge all extracted data into a final JSON object.
2. If the same column appears in multiple documents, prefer the most complete/detailed value.
3. Some columns are "linkage" fields that cross-reference between documents - ensure they match.
4. Return a JSON object with two keys:
   - "shared_row": columns that should be the same across all rows (Commercial Invoice C–O, Tax Invoice AR–BG)
   - "per_vehicle_rows": array of objects, one per BoE in Case 1 (or one per vehicle in other cases). Each object contains BoE-specific, EWB-specific, and transport fields for that row.
5. Use column identifiers as keys (e.g., "Column C (3)").
6. IMPORTANT: Each unique BoE number (Case 1) or vehicle number must appear exactly once in per_vehicle_rows. Merge ALL data for the same BoE/vehicle into a single row — do NOT collapse multiple BoEs into one row.
7. Preserve date strings (DD-MMM-YYYY) and INR-prefixed amounts exactly as extracted. Do not convert to numbers or datetime objects.

## Response:
Return ONLY a valid JSON object. No markdown, no explanation."""


def _format_dict(d: dict) -> str:
    import json
    return json.dumps(d, indent=2, default=str)


def _format_existing(rows: list[dict]) -> str:
    lines = []
    for i, row in enumerate(rows[:3]):  # Show first 3 rows as example
        filtered = {k: v for k, v in row.items() if k != "_row_num" and v is not None}
        lines.append(f"Row {i+1}: {_format_dict(filtered)}")
    if len(rows) > 3:
        lines.append(f"... and {len(rows) - 3} more rows")
    return "\n".join(lines)
