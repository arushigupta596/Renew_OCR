from __future__ import annotations

import json
import re
from openai import OpenAI
from parsers.prompts import (
    build_extraction_prompt,
    build_multi_row_extraction_prompt,
    build_merge_prompt,
)

# Document types that may contain per-vehicle/per-row data
MULTI_ROW_DOCS = {"e-way bill", "lorry receipt", "lr", "grn", "goods receipt"}


def _is_multi_row_doc(filename: str) -> bool:
    """Check if a document likely contains per-row data."""
    lower = filename.lower()
    return any(kw in lower for kw in ["e-way", "eway", "lr", "lorry", "grn", "goods receipt"])


def _clean_json_response(text: str) -> str:
    """Strip markdown code fences and extract JSON from LLM response."""
    text = text.strip()
    # Remove ```json ... ``` wrapping
    match = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if match:
        text = match.group(1).strip()
    return text


def parse_document(
    ocr_text: str,
    column_headers: list[str],
    filename: str,
    api_key: str,
    model: str,
    base_url: str,
) -> dict:
    """
    Send OCR text to LLM via OpenRouter and extract structured data.

    Returns a dict mapping column identifiers to extracted values.
    For multi-row documents, returns {"shared": {...}, "per_row": [...]}.
    """
    client = OpenAI(api_key=api_key, base_url=base_url)

    if _is_multi_row_doc(filename):
        prompt = build_multi_row_extraction_prompt(ocr_text, column_headers, filename)
    else:
        prompt = build_extraction_prompt(ocr_text, column_headers, filename)

    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=8000,
    )

    raw = response.choices[0].message.content
    cleaned = _clean_json_response(raw)

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return {"_raw_response": raw, "_parse_error": "Failed to parse JSON"}


def parse_all_documents(
    ocr_results: dict[str, str],
    column_headers: list[str],
    api_key: str,
    model: str,
    base_url: str,
    progress_callback=None,
) -> dict[str, dict]:
    """
    Parse all OCR results through the LLM.

    Returns dict mapping filename to extracted data.
    """
    all_extracted = {}
    total = len(ocr_results)

    for i, (filename, ocr_text) in enumerate(ocr_results.items()):
        if ocr_text.startswith("[OCR ERROR]"):
            all_extracted[filename] = {"_error": ocr_text}
        else:
            all_extracted[filename] = parse_document(
                ocr_text, column_headers, filename, api_key, model, base_url
            )

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return all_extracted


def merge_all_extractions(
    all_extracted: dict[str, dict],
    column_headers: list[str],
    existing_data: list[dict] | None,
    api_key: str,
    model: str,
    base_url: str,
) -> dict:
    """
    Merge extracted data from all documents into final row data.

    Returns {"shared_row": {...}, "per_vehicle_rows": [...]}.
    """
    client = OpenAI(api_key=api_key, base_url=base_url)

    prompt = build_merge_prompt(all_extracted, column_headers, existing_data)

    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=16000,
    )

    raw = response.choices[0].message.content
    cleaned = _clean_json_response(raw)

    try:
        result = json.loads(cleaned)
        return _deduplicate_per_vehicle_rows(result, column_headers)
    except json.JSONDecodeError:
        return {"_raw_response": raw, "_parse_error": "Failed to parse JSON"}


def _deduplicate_per_vehicle_rows(merged: dict, column_headers: list[str]) -> dict:
    """Remove duplicate per_vehicle_rows using invoice number as the unique key.

    Works across all case templates by searching row keys/values for
    invoice-related content rather than relying on exact column header format.
    Falls back to full-row content hash for exact duplicate removal.
    """
    per_vehicle = merged.get("per_vehicle_rows", [])
    if not per_vehicle:
        return merged

    def _find_invoice_value(row: dict) -> str | None:
        """Find the invoice number value in a row by checking all keys."""
        for key, val in row.items():
            key_lower = str(key).lower()
            if "invoice" in key_lower and ("no" in key_lower or "num" in key_lower or "number" in key_lower):
                v = str(val).strip()
                if v and v.lower() not in ("", "none", "n/a", "null"):
                    return v
        # Broader search: any key with just "invoice" that has a number-like value
        for key, val in row.items():
            key_lower = str(key).lower()
            if "invoice" in key_lower:
                v = str(val).strip()
                if v and v.lower() not in ("", "none", "n/a", "null"):
                    return v
        return None

    def _row_hash(row: dict) -> str:
        """Create a normalized hash of the entire row for exact-duplicate detection."""
        parts = []
        for k in sorted(row.keys()):
            v = str(row[k]).strip().upper()
            if v and v not in ("NONE", "N/A", "NULL", ""):
                parts.append(f"{k}={v}")
        return "|".join(parts)

    seen_invoices: set[str] = set()
    seen_hashes: set[str] = set()
    unique_rows: list[dict] = []

    for row in per_vehicle:
        invoice = _find_invoice_value(row)
        if invoice:
            norm = invoice.strip().upper()
            if norm in seen_invoices:
                continue
            seen_invoices.add(norm)
            unique_rows.append(row)
        else:
            # No invoice key found — fall back to full-row hash
            h = _row_hash(row)
            if h in seen_hashes:
                continue
            seen_hashes.add(h)
            unique_rows.append(row)

    merged["per_vehicle_rows"] = unique_rows
    return merged
