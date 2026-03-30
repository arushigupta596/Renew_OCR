def build_extraction_prompt(
    ocr_text: str,
    column_headers: list[str],
    doc_filename: str,
) -> str:
    """
    Build a prompt that asks the LLM to extract data from OCR text
    and map it to Excel column headers.
    """
    headers_str = "\n".join(column_headers)

    return f"""You are a document data extraction expert. You are given OCR-extracted text from a PDF document and a list of Excel column headers. Your job is to extract values from the document that correspond to the column headers.

## Document filename: {doc_filename}

## Excel Column Headers (extract values for these):
{headers_str}

## Rules:
1. Return a JSON object where keys are the column identifiers (e.g., "Column B (2)") and values are the extracted data.
2. Only include columns where you can find a corresponding value in the document. Skip columns that don't relate to this document.
3. For dates, use YYYY-MM-DD format.
4. For monetary amounts, return numbers without currency symbols or commas.
5. For Yes/No fields, return "Yes" or "No".
6. For text fields with addresses, preserve line breaks as \\n.
7. If a field is clearly present but the value is not applicable, use "N/A".
8. Do NOT guess or fabricate values. Only extract what is clearly stated in the document.

## OCR-Extracted Document Text:
{ocr_text}

## Response:
Return ONLY a valid JSON object. No markdown formatting, no explanation, just the JSON."""


def build_multi_row_extraction_prompt(
    ocr_text: str,
    column_headers: list[str],
    doc_filename: str,
) -> str:
    """
    Build a prompt for documents that may contain per-vehicle/per-item data
    (e.g., E-Way Bill vehicle list, LR copies, GRN with multiple entries).
    """
    headers_str = "\n".join(column_headers)

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
3. For dates, use YYYY-MM-DD format. For amounts, return numbers without currency symbols.
4. Do NOT guess or fabricate values.

## OCR-Extracted Document Text:
{ocr_text}

## Response:
Return ONLY a valid JSON object with "shared" and "per_row" keys. No markdown, no explanation."""


def build_merge_prompt(
    all_extracted: dict[str, dict],
    column_headers: list[str],
    existing_row_data: list[dict] | None = None,
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

    return f"""You are a data reconciliation expert. You have extracted data from multiple PDF documents that all belong to the same shipment/transaction. Merge them into a complete dataset for the Excel tracker.

## Excel Column Headers:
{headers_str}

## Extracted Data from Each Document:
{docs_str}
{existing_str}

## Rules:
1. Merge all extracted data into a final JSON object.
2. If the same column appears in multiple documents, prefer the most complete/detailed value.
3. Some columns are "linkage" fields that cross-reference between documents - ensure they match.
4. Return a JSON object with two keys:
   - "shared_row": columns that should be the same across all rows (procurement, customs, e-way bill data)
   - "per_vehicle_rows": array of objects for per-vehicle data (transport, delivery), each with a "vehicle_no" key for matching
5. Use column identifiers as keys (e.g., "Column B (2)").

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
