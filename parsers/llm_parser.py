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
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return {"_raw_response": raw, "_parse_error": "Failed to parse JSON"}
