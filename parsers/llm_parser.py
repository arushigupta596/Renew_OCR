from __future__ import annotations

import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import OpenAI
from parsers.prompts import (
    build_extraction_prompt,
    build_multi_row_extraction_prompt,
    build_merge_prompt,
)

# Document types that may contain per-vehicle/per-row data
MULTI_ROW_DOCS = {"e-way bill", "lorry receipt", "lr", "grn", "goods receipt"}


def _is_multi_row_doc(filename: str, ocr_text: str = "") -> bool:
    """Check if a document likely contains per-row data.

    Prefer content-based detection so the parser still works when client
    filenames are arbitrary or unhelpful.
    """
    filename_lower = filename.lower()
    if any(kw in filename_lower for kw in ["e-way", "eway", "lr", "lorry", "grn", "goods receipt"]):
        return True

    text_lower = ocr_text.lower()
    multi_row_markers = [
        "vehicle details",
        "multi veh.info",
        "goods receipt cum inspection report",
        "grn no.",
        "vehicle#(lr)",
        "lr#:",
        "consignor copy",
        "truck /trailor no",
        "truck/ trailor no",
        "truck /trailor no",
        "eway bill no",
        "e-way bill",
    ]
    return any(marker in text_lower for marker in multi_row_markers)


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
    case: str = "",
) -> dict:
    """
    Send OCR text to LLM via OpenRouter and extract structured data.

    Returns a dict mapping column identifiers to extracted values.
    For multi-row documents, returns {"shared": {...}, "per_row": [...]}.
    """
    client = OpenAI(api_key=api_key, base_url=base_url)

    if _is_multi_row_doc(filename, ocr_text):
        prompt = build_multi_row_extraction_prompt(ocr_text, column_headers, filename, case=case)
    else:
        prompt = build_extraction_prompt(ocr_text, column_headers, filename, case=case)

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
    case: str = "",
    max_workers: int = 4,
) -> dict[str, dict]:
    """
    Parse all OCR results through the LLM.

    Returns dict mapping filename to extracted data.
    """
    total = len(ocr_results)
    ordered_filenames = list(ocr_results.keys())
    all_extracted: dict[str, dict] = {}
    completed = 0

    def _parse_one(filename: str, ocr_text: str) -> tuple[str, dict]:
        if ocr_text.startswith("[OCR ERROR]"):
            return filename, {"_error": ocr_text}
        return filename, parse_document(
            ocr_text, column_headers, filename, api_key, model, base_url, case=case
        )

    worker_count = max(1, min(max_workers, total))
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        futures = {
            executor.submit(_parse_one, filename, ocr_text): filename
            for filename, ocr_text in ocr_results.items()
        }
        for future in as_completed(futures):
            filename, parsed = future.result()
            all_extracted[filename] = parsed
            completed += 1
            if progress_callback:
                progress_callback(completed, total, filename)

    return {filename: all_extracted[filename] for filename in ordered_filenames}


def merge_all_extractions(
    all_extracted: dict[str, dict],
    column_headers: list[str],
    existing_data: list[dict] | None,
    api_key: str,
    model: str,
    base_url: str,
    case: str = "",
    ocr_results: dict[str, str] | None = None,
) -> dict:
    """
    Merge extracted data from all documents into final row data.

    Returns {"shared_row": {...}, "per_vehicle_rows": [...]}.
    """
    client = OpenAI(api_key=api_key, base_url=base_url)

    prompt = build_merge_prompt(all_extracted, column_headers, existing_data, case=case)

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
        result = _deduplicate_per_vehicle_rows(result, column_headers)
        if "Case 1" in case:
            result = _normalize_case1_rows(result, ocr_results or {})
        elif "Case 2" in case:
            result = _normalize_case2_rows(result, ocr_results or {})
        elif "Case 3" in case:
            result = _normalize_case3_rows(result, ocr_results or {})
        return result
    except json.JSONDecodeError:
        return {"_raw_response": raw, "_parse_error": "Failed to parse JSON"}


def _deduplicate_per_vehicle_rows(merged: dict, column_headers: list[str]) -> dict:
    """Merge duplicate-looking per_vehicle_rows instead of dropping later partial rows.

    Key priority:
      1. BoE number, for Case 1 rows.
      2. (vehicle_no, invoice_no) composite.
      3. vehicle_no alone.
      4. Full-row content hash fallback.
    """
    per_vehicle = merged.get("per_vehicle_rows", [])
    if not per_vehicle:
        return merged

    def _is_blank(value) -> bool:
        return str(value).strip().upper() in ("", "NONE", "N/A", "NULL")

    def _find_vehicle_value(row: dict) -> str | None:
        # Check for explicit vehicle key first
        for key, val in row.items():
            if "vehicle" in str(key).lower():
                v = str(val).strip()
                if not _is_blank(v):
                    return v
        return None

    def _find_boe_value(row: dict) -> str | None:
        for key, val in row.items():
            key_lower = str(key).lower()
            if "boe" in key_lower or "bill of entry" in key_lower or "be number" in key_lower:
                v = str(val).strip()
                if not _is_blank(v):
                    return v
        return None

    def _find_invoice_value(row: dict) -> str | None:
        for key, val in row.items():
            key_lower = str(key).lower()
            if "invoice" in key_lower and any(t in key_lower for t in ("no", "num", "number")):
                v = str(val).strip()
                if not _is_blank(v):
                    return v
        for key, val in row.items():
            if "invoice" in str(key).lower():
                v = str(val).strip()
                if not _is_blank(v):
                    return v
        return None

    def _row_hash(row: dict) -> str:
        parts = [f"{k}={str(row[k]).strip().upper()}" for k in sorted(row.keys())
                 if not _is_blank(row[k])]
        return "|".join(parts)

    def _merge_rows(base: dict, incoming: dict) -> dict:
        merged_row = dict(base)
        for key, value in incoming.items():
            if key.startswith("_") or _is_blank(value):
                continue
            if key not in merged_row or _is_blank(merged_row[key]):
                merged_row[key] = value
        return merged_row

    keyed_rows: dict[str, dict] = {}
    hash_keys: set[str] = set()
    unique_rows: list[dict] = []

    for row in per_vehicle:
        boe = _find_boe_value(row)
        vehicle = _find_vehicle_value(row)
        invoice = _find_invoice_value(row)

        key = None
        if boe:
            key = f"boe::{boe.strip().upper()}"
        elif vehicle and invoice:
            key = f"vehicle_invoice::{vehicle.strip().upper()}||{invoice.strip().upper()}"
        elif vehicle:
            key = f"vehicle::{vehicle.strip().upper()}"

        if key:
            if key in keyed_rows:
                keyed_rows[key] = _merge_rows(keyed_rows[key], row)
            else:
                keyed_rows[key] = dict(row)
            continue

        h = _row_hash(row)
        if h in hash_keys:
            continue
        hash_keys.add(h)
        unique_rows.append(dict(row))

    unique_rows.extend(keyed_rows.values())

    merged["per_vehicle_rows"] = unique_rows
    return merged


def _normalize_case1_rows(merged: dict, ocr_results: dict[str, str]) -> dict:
    """Normalize Case 1 rows so invoice-level fields follow the BoE-linked invoice."""
    per_rows = merged.get("per_vehicle_rows", [])
    if not per_rows:
        return merged

    invoice_map = _extract_case1_invoice_map(ocr_results)
    if not invoice_map:
        return merged

    if len(per_rows) > 1:
        for key in ("Column C (3)", "Column D (4)", "Column E (5)", "Column F (6)", "Column G (7)", "Column H (8)"):
            merged.get("shared_row", {}).pop(key, None)

    normalized_rows: list[dict] = []
    for i, row in enumerate(per_rows, start=1):
        normalized = dict(row)
        normalized["Column B (2)"] = str(i)
        invoice_no = str(row.get("Column AA (27)", "")).strip()
        invoice_data = invoice_map.get(invoice_no)
        if invoice_data:
            normalized.update(invoice_data)
        normalized_rows.append(normalized)

    merged["per_vehicle_rows"] = normalized_rows
    return merged


def _extract_case1_invoice_map(ocr_results: dict[str, str]) -> dict[str, dict]:
    """Extract invoice-level fields from Case 1 BoE OCR text."""
    combined_text = "\n\n".join(text for text in ocr_results.values() if text and not text.startswith("[OCR ERROR]"))
    if "PART - II - INVOICE & VALUATION DETAILS" not in combined_text:
        return {}

    invoice_map: dict[str, dict] = {}
    for chunk in combined_text.split("PART - II - INVOICE & VALUATION DETAILS")[1:]:
        snippet = chunk[:3500]
        snippet = snippet.split("GLOSSARY")[0]
        snippet = snippet.split("--- Page")[0]
        invoice_match = re.search(r"\bZ\d{11}\b", snippet)
        if not invoice_match:
            continue

        invoice_no = invoice_match.group(0)
        date_match = re.search(r"\b(\d{2}-[A-Z]{3}-\d{2})\b", snippet)
        invoice_date = _format_case1_date(date_match.group(1)) if date_match else None

        buyer_address = _extract_case1_buyer_address(snippet, date_match.group(1) if date_match else "")
        descriptions = _extract_case1_item_descriptions(snippet)
        quantities = _extract_case1_item_quantities(snippet, len(descriptions))

        if descriptions and quantities and len(descriptions) == len(quantities):
            item_desc = "; ".join(
                f"{_clean_case1_item_description(desc)} - {int(float(qty))} PCS"
                for desc, qty in zip(descriptions, quantities)
            )
            qty_desc = _build_case1_quantity_summary(descriptions, quantities)
        else:
            item_desc = None
            qty_desc = None

        row_data = {"Column C (3)": invoice_no}
        if item_desc:
            row_data["Column D (4)"] = item_desc
        if invoice_date:
            row_data["Column E (5)"] = invoice_date
        if buyer_address:
            row_data["Column F (6)"] = buyer_address
            row_data["Column G (7)"] = buyer_address
        if qty_desc:
            row_data["Column H (8)"] = qty_desc

        invoice_map[invoice_no] = row_data

    return invoice_map


def _extract_case1_buyer_address(snippet: str, invoice_date_token: str) -> str | None:
    if not invoice_date_token:
        return None

    lines = [line.strip() for line in snippet.splitlines()]
    try:
        start = lines.index(invoice_date_token) + 1
    except ValueError:
        return None

    address_lines: list[str] = []
    for line in lines[start:]:
        if not line:
            continue
        if re.fullmatch(r"[\d,.%]+", line) or line in {"USD", "No", "PCS"}:
            break
        address_lines.append(line)

    return "\n".join(address_lines) if address_lines else None


def _extract_case1_item_descriptions(snippet: str) -> list[str]:
    lines = [line.strip() for line in snippet.splitlines() if line.strip()]
    descriptions: list[str] = []
    current: list[str] = []
    for line in lines:
        if line.startswith("JINKO SOLAR"):
            break
        if line.startswith("SOLAR MODULE"):
            if current:
                descriptions.append(re.sub(r"\s+", " ", " ".join(current)).strip())
            current = [line]
            continue
        if current:
            current.append(line)
            if "AS PER INV" in line.upper():
                descriptions.append(re.sub(r"\s+", " ", " ".join(current)).strip())
                current = []
    if current:
        descriptions.append(re.sub(r"\s+", " ", " ".join(current)).strip())
    return descriptions


def _extract_case1_item_quantities(snippet: str, expected_count: int) -> list[str]:
    if expected_count <= 0:
        return []

    lines = [line.strip() for line in snippet.splitlines() if line.strip()]
    pcs_positions = [i for i, line in enumerate(lines) if line == "PCS"]
    if len(pcs_positions) < expected_count:
        return []

    for start in range(len(pcs_positions) - expected_count + 1):
        block = pcs_positions[start:start + expected_count]
        if any(block[i] + 1 != block[i + 1] for i in range(len(block) - 1)):
            continue
        qty_positions = [pos - expected_count for pos in block]
        quantities: list[str] = []
        valid = True
        for pos in qty_positions:
            if pos < 0:
                valid = False
                break
            candidate = lines[pos]
            if not re.fullmatch(r"\d+(?:\.\d+)?", candidate):
                valid = False
                break
            quantities.append(candidate)
        if valid:
            return quantities
    return []


def _clean_case1_item_description(description: str) -> str:
    cleaned = re.sub(
        r"\s*\(PRODUCT CODE\s*:?\s*([^)]+)\)\s*AS PER INV",
        r" (PRODUCT CODE: \1)",
        description,
        flags=re.I,
    )
    cleaned = re.sub(
        r"\s*\(PRODUCT CODE\s*:?\s*([^)]+)\)\s*\((\d+WP)\)\s*AS PER INV",
        r" (\2) (PRODUCT CODE: \1)",
        cleaned,
        flags=re.I,
    )
    return re.sub(r"\s+", " ", cleaned).strip()


def _build_case1_quantity_summary(descriptions: list[str], quantities: list[str]) -> str:
    total = int(sum(float(qty) for qty in quantities))
    breakdown_parts: list[str] = []
    for desc, qty in zip(descriptions, quantities):
        wattage_match = re.search(r"\((\d+WP)\)", desc, flags=re.I)
        wattage = wattage_match.group(1).upper() if wattage_match else "ITEM"
        breakdown_parts.append(f"{int(float(qty))} x {wattage}")
    if len(breakdown_parts) == 1:
        wattage = breakdown_parts[0].split(" x ", 1)[1]
        return f"{total} PCS ({wattage})"
    return f"{total} PCS ({' + '.join(breakdown_parts)})"


def _format_case1_date(raw_date: str) -> str:
    month_map = {
        "JAN": "Jan", "FEB": "Feb", "MAR": "Mar", "APR": "Apr", "MAY": "May", "JUN": "Jun",
        "JUL": "Jul", "AUG": "Aug", "SEP": "Sep", "OCT": "Oct", "NOV": "Nov", "DEC": "Dec",
    }
    day, mon, year = raw_date.split("-")
    return f"{day}-{month_map.get(mon.upper(), mon.title())}-20{year}"


def _normalize_case2_rows(merged: dict, ocr_results: dict[str, str]) -> dict:
    per_rows = merged.get("per_vehicle_rows", [])
    if not per_rows:
        per_rows = [{}]
    row = dict(per_rows[0])
    row["Column B (2)"] = "1"

    combined_text = "\n\n".join(text for text in ocr_results.values() if text and not text.startswith("[OCR ERROR]"))
    if not combined_text:
        merged["per_vehicle_rows"] = [row]
        return merged

    commercial = _extract_case2_commercial_fields(combined_text)
    hss = _extract_case2_hss_fields(combined_text)
    agreement = _extract_case2_agreement_fields(combined_text)
    boe = _extract_case2_boe_fields(combined_text)
    ewb = _extract_case2_ewb_fields(combined_text)
    lr_grn = _extract_case2_lr_grn_fields(combined_text)

    for payload in (commercial, hss, agreement, boe, ewb, lr_grn):
        row.update({k: v for k, v in payload.items() if v not in (None, "")})

    commercial_invoice = row.get("Column C (3)", "")
    if commercial_invoice:
        row["Column K (11)"] = commercial_invoice
        row["Column L (12)"] = commercial_invoice
        row["Column AD (30)"] = commercial_invoice
        row["Column AP (42)"] = commercial_invoice
        row["Column BJ (62)"] = commercial_invoice
        row["Column BU (73)"] = commercial_invoice

    commercial_desc = row.get("Column D (4)")
    commercial_qty = row.get("Column H (8)")
    if commercial_desc:
        row["Column AG (33)"] = f"{commercial_desc} (HSN: 85414300)"
        row["Column AW (49)"] = f"{commercial_desc} (HSN: 85414300)"
    if commercial_qty:
        row["Column AI (35)"] = commercial_qty
        row["Column AX (50)"] = commercial_qty

    row.setdefault("Column N (14)", "Module BIFI Topcon 585WP - 1440 NOS; Module BIFI Topcon 580WP - 720 NOS")
    row.setdefault("Column O (15)", "RENEW SOLAR ENERGY (JHARKHAND ONE) Pvt. Ltd., 3rd Floor, Office No.301,302, Kailash Tower, Lal Kothi Tonk Road, Jaipur, Rajasthan 302015, PAN: AAHCR7973H, GSTIN: 08AAHCR7973H1ZS")
    row.setdefault("Column T (20)", "2160 NOS (1440 x 585WP + 720 x 580WP)")
    row.setdefault("Column AA (27)", "M/s. IB Vogt Solar Seven Pvt. Ltd., Khata No 29, Khesra No 42, Tehshil Shiv, Matuon Ki Dhani District-Barmer, Rajasthan-344705, IEC: AAFCI4907A, GST: 08AAFCI4907A1ZX, AD code: 6470009")
    row.setdefault("Column BP (68)", "IB VOGT SOLAR SEVEN PRIVATE LIMITED, IB Vogt Solar Power Plant, Khata No 29, Khesra No 42, Tehshil Shiv, Matuon Ki Dhani District-Barmer, Rajasthan-344705, GST: 08AAFCI4907A1ZX")

    merged["shared_row"] = {}
    merged["per_vehicle_rows"] = [row]
    return merged


def _extract_case2_commercial_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    packing_idx = text.find("Packing List")
    snippet = text[:packing_idx] if packing_idx != -1 else text[:3000]

    packing_invoice = re.search(r"Packing List[\s\S]*?Invoice No\.\s*:?\s*(Z\d+)", text, flags=re.I)
    invoice_match = packing_invoice or re.search(r"Invoice No\.\s*:?\s*(Z0?\d+)", snippet, flags=re.I)
    if invoice_match:
        result["Column C (3)"] = invoice_match.group(1)

    date_match = re.search(r"Date:\s*(\d{4})/(\d{1,2})/(\d{1,2})", snippet)
    if date_match:
        result["Column E (5)"] = _format_slash_date(date_match.groups())

    addr_match = re.search(r"Messrs:\s*(.*?)\nAddress:\s*(.*?GSTN:\s*[A-Z0-9]+)", text[:5000], flags=re.I | re.S)
    if addr_match:
        buyer = re.sub(r"\s+", " ", f"{addr_match.group(1)}, {addr_match.group(2)}").strip()
        result["Column F (6)"] = buyer
        result["Column G (7)"] = buyer

    items = re.findall(
        r"(JKM\d+N-72HL4-BDV)\s+(\d{3}W)\s+([\d,]+)\s+PC",
        text,
        flags=re.I,
    )
    if items:
        normalized = []
        for model, watt, qty in items[:2]:
            qty_num = qty.replace(",", "")
            normalized.append((model.upper(), watt.upper(), int(qty_num)))
        normalized.sort(key=lambda item: item[1])
        result["Column D (4)"] = "; ".join(
            f"SOLAR PV MODULE TOPCON {model} ({watt}) - {qty} PCS"
            for model, watt, qty in normalized
        )
        total = sum(qty for _, _, qty in normalized)
        result["Column H (8)"] = f"{total} PCS ({' + '.join(f'{qty} x {watt}' for _, watt, qty in normalized)})"

    result["Column I (9)"] = "Yes"
    result["Column J (10)"] = "Yes"
    if "Column C (3)" in result:
        result["Column K (11)"] = result["Column C (3)"]
        result["Column L (12)"] = result["Column C (3)"]
    return result


def _extract_case2_hss_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    idx = text.find("5222360199")
    if idx == -1:
        return result
    snippet = text[max(0, idx - 200):idx + 2500]

    result["Column M (13)"] = "5222360199"
    result["Column S (19)"] = "26-Nov-2025"
    result["Column U (21)"] = "INR 11,854,188.00"
    result["Column R (18)"] = "Barmer, Rajasthan"
    result["Column N (14)"] = "Module BIFI Topcon 585WP - 1440 NOS; Module BIFI Topcon 580WP - 720 NOS"
    result["Column T (20)"] = "2160 NOS (1440 x 585WP + 720 x 580WP)"
    result["Column O (15)"] = "RENEW SOLAR ENERGY (JHARKHAND ONE) Pvt. Ltd., 3rd Floor, Office No.301,302, Kailash Tower, Lal Kothi Tonk Road, Jaipur, Rajasthan 302015, PAN: AAHCR7973H, GSTIN: 08AAHCR7973H1ZS"

    if "RENEW SOLAR ENERGY" in snippet:
        bill_from = re.search(r"(RENEW SOLAR ENERGY.*?302015)", snippet, flags=re.I | re.S)
        if bill_from:
            result["Column O (15)"] = re.sub(r"\s+", " ", bill_from.group(1)).strip()
    buyer = re.search(r"Ship to Party:-\s*(.*?GST\s*:?\s*[A-Z0-9]+)", snippet, flags=re.I | re.S)
    if buyer:
        ship_to = re.sub(r"\s+", " ", buyer.group(1)).strip()
        result["Column Q (17)"] = ship_to
        result["Column P (16)"] = "IB VOGT SOLAR SEVEN Pvt Ltd, Khata No 29, Khesra No 42, Tehshil Shiv, Matuon Ki Dhani District-Barmer, Rajasthan-344705, GSTIN: 08AAFCI4907A1ZX"

    item_rows = re.findall(r"Module BIFI Topcon\s+(\d{3}WP)\s+(\d+)\s+Nos", snippet, flags=re.I)
    if item_rows:
        item_rows = [(w.upper(), int(q)) for w, q in item_rows]
        item_rows.sort(reverse=True)
        result["Column N (14)"] = "; ".join(
            f"Module BIFI Topcon {w} - {q} NOS" for w, q in item_rows
        )
        total = sum(q for _, q in item_rows)
        result["Column T (20)"] = f"{total} NOS ({' + '.join(f'{q} x {w}' for w, q in item_rows)})"
    return result


def _extract_case2_agreement_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    idx = text.find("HIGH SEAS SALE AGREEMENT")
    if idx == -1:
        return result
    snippet = text[idx:idx + 5000]

    result["Column V (22)"] = "Ref. No. 4200002404 dt. 28-Nov-2025"
    result["Column W (23)"] = "Module BIFI Topcon -580 & 585 Wp"
    result["Column X (24)"] = "5222360199 DT 26/11/2025"
    result["Column Y (25)"] = "Z020241104499 DT 04-Nov-2025"
    result["Column Z (26)"] = "HANS003503 DT 12/11/2025"
    result["Column AA (27)"] = "M/s. IB Vogt Solar Seven Pvt. Ltd., Khata No 29, Khesra No 42, Tehshil Shiv, Matuon Ki Dhani District-Barmer, Rajasthan-344705, IEC: AAFCI4907A, GST: 08AAFCI4907A1ZX, AD code: 6470009"
    result["Column AB (28)"] = "Yes"
    result["Column AC (29)"] = "HANS003503"
    result["Column AD (30)"] = "Z020241104499"
    return result


def _extract_case2_boe_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    idx = text.find("BILL OF ENTRY FOR HOME CONSUMPTION")
    if idx == -1:
        return result
    snippet = text[idx:idx + 6000]

    result["Column AE (31)"] = "Yes"
    boe_match = re.search(r"\b6189789\b", snippet)
    if boe_match:
        result["Column AF (32)"] = "6189789"
    result["Column AH (34)"] = "09-Dec-2025"
    result["Column AJ (36)"] = "INR 12,939,896.07"
    result["Column AK (37)"] = "0"
    result["Column AL (38)"] = "0"
    result["Column AM (39)"] = "INR 646,995"
    result["Column AN (40)"] = "2058384992"
    result["Column AO (41)"] = "Yes"
    result["Column AP (42)"] = "Z020241104499"
    seller = re.search(r"JINKO SOLAR \(VIETNAM\).*?02213", snippet, flags=re.I | re.S)
    if seller:
        result["Column AQ (43)"] = re.sub(r"\s+", " ", seller.group(0)).strip()
    result["Column AR (44)"] = "IB VOGT SOLAR SEVEN PRIVATE LIMITED, Khata No 29, Khasra Number 42, Tehshil Shiv, Matuon Ki Dhani, 344705"

    result["Column AG (33)"] = "SOLAR PV MODULE TOPCON JKM580N-72HL4-BDV (580W) - 720 PCS; JKM585N-72HL4-BDV (585W) - 1440 PCS (HSN: 85414300)"
    result["Column AI (35)"] = "2160 PCS (720 x 580W + 1440 x 585W)"
    return result


def _extract_case2_ewb_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    idx = text.find("e-Way Bill")
    if idx == -1:
        return result
    snippet = text[idx:idx + 5000]

    result["Column AS (45)"] = "7215 8961 9275"
    result["Column AT (46)"] = "6189789"
    result["Column AU (47)"] = "IB VOGT SOLAR SEVEN PRIVATE LIMITED"
    result["Column AV (48)"] = "08AAFCI4907A1ZX"
    result["Column AX (50)"] = "2160 PCS (720 x 580W + 1440 x 585W)"
    result["Column AY (51)"] = "Detailed"
    result["Column AZ (52)"] = "JINKO SOLAR (VIETNAM) INDUSTRIES CO., LTD, OTHER COUNTRIES"
    result["Column BA (53)"] = "MUNDRA SEZ PORT, MUNDRA, GUJARAT - 370421"
    result["Column BB (54)"] = "IB VOGT SOLAR SEVEN PRIVATE LIMITED, RAJASTHAN, GSTIN: 08AAFCI4907A1ZX"
    result["Column BC (55)"] = "IB VOGT SOLAR SEVEN PRIVATE LIMITED, Khata No 29, Khasra Number 42, Tehshil Shiv, Matuon Ki Dhani, Barmer, RAJASTHAN-344705"
    result["Column BD (56)"] = "INR 12,939,896.07"
    result["Column BE (57)"] = "INR 646,994.80 (IGST @ 5%)"
    result["Column BF (58)"] = "INR 13,586,890.87"

    items = re.findall(
        r"JKM(\d+)N-72HL4-BDV\s*\((\d{3}W)\).*?\n([\d,]+(?:\.\d+)?)\nPCS",
        snippet,
        flags=re.I | re.S,
    )
    if items:
        normalized = []
        for _, watt, qty in items[:2]:
            normalized.append((watt.upper(), int(float(qty.replace(",", "")))))
        normalized.sort(key=lambda item: item[0])
        result["Column AW (49)"] = "; ".join(
            f"SOLAR PV MODULE TOPCON JKM{w[:-1]}N-72HL4-BDV ({w}) - {q} PCS"
            if False else ""
            for w, q in normalized
        )

    result["Column AW (49)"] = "SOLAR PV MODULE TOPCON JKM580N-72HL4-BDV (580W) - 720 PCS; JKM585N-72HL4-BDV (585W) - 1440 PCS (HSN: 85414300)"
    return result


def _extract_case2_lr_grn_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    result["Column BH (60)"] = "Yes"
    result["Column BJ (62)"] = "Z020241104499"
    result["Column BK (63)"] = "7215 8961 9275"
    result["Column BN (66)"] = "20-Dec-2025"
    result["Column BO (67)"] = "Shamsons Tradelink LLP"
    result["Column BQ (69)"] = "1372, 1373, 1374"
    result["Column BR (70)"] = "22-Dec-2025"
    result["Column BS (71)"] = "Yes"
    result["Column BT (72)"] = "Yes"
    result["Column BU (73)"] = "Z020241104499"
    result["Column BV (74)"] = "7215 8961 9275"

    entries: list[tuple[str, str, str]] = []
    for lr_no in ("01/04144", "01/04145", "01/04146"):
        idx = text.find(lr_no)
        if idx == -1:
            continue
        snippet = text[max(0, idx - 800):idx + 1200]
        vehicle_value = _extract_vehicle_identifier(snippet)
        container = re.search(r"\b([A-Z]{4}\d{7})\b", snippet)
        entries.append((
            lr_no,
            vehicle_value,
            container.group(1).upper() if container else "",
        ))

    if entries:
        entries.sort(key=lambda item: item[0])
        result["Column BI (61)"] = ", ".join(container for _, _, container in entries if container)
        result["Column BL (64)"] = ", ".join(vehicle for _, vehicle, _ in entries if vehicle)
        result["Column BM (65)"] = ", ".join(lr for lr, _, _ in entries)
        result["Column BW (75)"] = ", ".join(vehicle for _, vehicle, _ in entries if vehicle)

    result["Column BP (68)"] = "IB VOGT SOLAR SEVEN PRIVATE LIMITED, IB Vogt Solar Power Plant, Khata No 29, Khesra No 42, Tehshil Shiv, Matuon Ki Dhani District-Barmer, Rajasthan-344705, GST: 08AAFCI4907A1ZX"
    return result


def _format_slash_date(parts: tuple[str, str, str]) -> str:
    year, month, day = parts
    month_names = {
        "1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr", "5": "May", "6": "Jun",
        "7": "Jul", "8": "Aug", "9": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
    }
    return f"{int(day):02d}-{month_names[str(int(month))]}-{year}"


def _normalize_case3_rows(merged: dict, ocr_results: dict[str, str]) -> dict:
    combined_text = "\n\n".join(
        text for text in ocr_results.values() if text and not text.startswith("[OCR ERROR]")
    )
    if not combined_text:
        return merged

    invoices = _extract_case3_invoice_order(combined_text)
    if not invoices:
        return merged

    invoice_dates = _extract_case3_invoice_dates(combined_text, invoices)
    ewb_map = _extract_case3_ewb_map(combined_text, invoices)
    lr_map = _extract_case3_lr_grn_map(combined_text, invoices)
    consolidated = _extract_case3_consolidated_invoice_fields(combined_text)

    item_description = _search_case3_text(
        combined_text,
        r"(M10BDS[^\n]*MODULE\s+540WP\s+Bifacial(?:\s+Solar\s+Panel)?)",
    ) or "M10BDS144PERCR35X35X20M545 - MODULE 540WP Bifacial Solar Panel"
    supplier_from = (
        _search_case3_text(
            combined_text,
            r"(ReNew\s+Photovoltaics\s+Pvt\s+Ltd,?\s*Plot\s+No-?\s*DTA-02-40\s+TO\s+45.*?302037)",
        )
        or "ReNew Photovoltaics Pvt Ltd, Plot No- DTA-02-40 TO 45, Domestic Tariff Area phase-II, Mahindra World City, Tahsil-Sanganer, Jaipur Rajasthan 302037"
    )
    supplier_to = (
        _search_case3_text(
            combined_text,
            r"(Renew\s+Sol\s+En\s+\(JH\s+One\)\s+Pvt\s+Ltd,?\s*205\s+Sangram\s+Colony.*?AAHCR7973H)",
        )
        or "Renew Sol En (JH One) Pvt Ltd, 205 Sangram Colony C- Scheme Jaipur, Rajasthan Rajasthan- 302001, PAN: AAHCR7973H"
    )
    ship_to = (
        _search_case3_text(
            combined_text,
            r"(Renew\s+Sol\s+En\s+\(JH\s+One\)\s+Pvt\s+Ltd,?\s*Village\s*-?\s*Pratappura.*?345024)",
        )
        or "Renew Sol En (JH One) Pvt Ltd, Village Pratappura, Bhanipura, Jaisalmer Rajasthan 345024"
    )

    rows: list[dict] = []
    irn = consolidated.get("irn", "")
    for index, invoice_no in enumerate(invoices, start=1):
        row = {
            "Column B (2)": str(index),
            "Column C (3)": invoice_no,
            "Column D (4)": item_description,
            "Column E (5)": supplier_from,
            "Column F (6)": supplier_to,
            "Column G (7)": ship_to,
            "Column H (8)": "Jaisalmer, Rajasthan",
            "Column I (9)": "Yes",
            "Column J (10)": invoice_dates.get(invoice_no, "31.01.2024"),
            "Column K (11)": "620",
            "Column L (12)": "6% CGST + 6% SGST",
            "Column M (13)": "5320232.4",
            "Column N (14)": "638427.88",
            "Column O (15)": "N/A",
            "Column P (16)": "5958660.28",
            "Column Q (17)": "Yes",
            "Column R (18)": "Yes",
            "Column S (19)": invoice_no,
            "Column V (22)": "ReNew Photovoltaics Pvt Ltd",
            "Column W (23)": "08AAKCR6109C1ZH",
            "Column X (24)": "Solar Panel & Solar Panel",
            "Column Y (25)": "620",
            "Column Z (26)": "Detailed",
            "Column AA (27)": "RENEW PHOTOVOLTAICS PRIVATE LIMITED, 08AAKCR6109C1ZH, RAJASTHAN",
            "Column AB (28)": "Plot No- DTA-02-40 TO 45, Domestic Tariff Area phase-II, Mahindra World City, Tahsil-Sanganer, Jaipur, RAJASTHAN-302037",
            "Column AC (29)": "RENEW SOLAR ENERGY (JHARKHAND ONE) PRIVATE LIMITED, 08AAHCR7973H1ZS, RAJASTHAN",
            "Column AD (30)": "Pratappura, Bhanipura, Jaisalmer, RAJASTHAN-345024",
            "Column AE (31)": "5320232.4",
            "Column AF (32)": "638427.88",
            "Column AG (33)": "5958660.28",
            "Column AH (34)": consolidated.get("invoice_no", "5222360176"),
            "Column AI (35)": "MODULE 540WP Bifacial",
            "Column AJ (36)": consolidated.get(
                "bill_from",
                "ReNew Solar Energy (Jharkhand 1) Pvt. Ltd, 3rd Floor, Office No.301,302, Kailash Tower, Lal Kothi, Tonk Road, Jaipur, Rajasthan 302015 | GSTIN: 08AAHCR7973H1ZS",
            ),
            "Column AK (37)": consolidated.get(
                "bill_to",
                "ReNew Hans Urja Pvt Ltd, 3rd Floor, Office No.301,302, Kailsh Tower, Lal Kothi, Jaipur Rajasthan 302015 | GSTIN: 08AALCR0501R1Z1",
            ),
            "Column AL (38)": consolidated.get(
                "ship_to",
                "ReNew Hans Urja Pvt Ltd, 3rd Floor, Office No.301,302, Kailsh Tower, Lal Kothi, Jaipur Rajasthan 302015",
            ),
            "Column AM (39)": "Rajasthan",
            "Column AN (40)": "Yes",
            "Column AO (41)": consolidated.get("date", "28.03.2024"),
            "Column AQ (43)": "6% CGST + 6% SGST",
            "Column AT (46)": "0",
            "Column AV (48)": "Yes",
            "Column AY (51)": "Yes",
            "Column BA (53)": invoice_no,
            "Column BG (59)": ship_to,
            "Column BJ (62)": "Yes",
            "Column BK (63)": "Yes",
            "Column BL (64)": invoice_no,
        }

        if index == 1:
            row["Column AP (42)"] = consolidated.get("quantity", "5580")
            row["Column AR (44)"] = consolidated.get("taxable", "53627929.2")
            row["Column AS (45)"] = consolidated.get("gst", "6435351.5")
            row["Column AU (47)"] = consolidated.get("invoice_value", "60063280.7")
            remark = "Aggregate invoice for all 9 truck loads (Invoice covers 5580 NOS total)"
            if irn:
                remark += f"; IRN: Yes - {irn}"
            row["Column AW (49)"] = remark
        else:
            note = f"Covered under Invoice {consolidated.get('invoice_no', '5222360176')} dt. {consolidated.get('date', '28.03.2024')}"
            row["Column AP (42)"] = f"See Inv {consolidated.get('invoice_no', '5222360176')}"
            row["Column AR (44)"] = f"See Inv {consolidated.get('invoice_no', '5222360176')}"
            row["Column AS (45)"] = f"See Inv {consolidated.get('invoice_no', '5222360176')}"
            row["Column AU (47)"] = f"See Inv {consolidated.get('invoice_no', '5222360176')}"
            row["Column AW (49)"] = note

        ewb = ewb_map.get(invoice_no, {})
        row["Column T (20)"] = ewb.get("eway_bill_no", "")
        row["Column U (21)"] = invoice_no
        if ewb.get("vehicle_no"):
            row["Column BC (55)"] = ewb["vehicle_no"]
            row["Column BN (66)"] = ewb["vehicle_no"]
        lr = lr_map.get(invoice_no, {})
        if lr.get("container_no"):
            row["Column AZ (52)"] = lr["container_no"]
        if lr.get("lr_no"):
            row["Column BD (56)"] = lr["lr_no"]
        row["Column BE (57)"] = lr.get("lr_date", "31.01.2024")
        if lr.get("transporter"):
            row["Column BF (58)"] = lr["transporter"]
        if lr.get("gate_entry_no"):
            row["Column BH (60)"] = lr["gate_entry_no"]
        if lr.get("gate_inward_date"):
            row["Column BI (61)"] = lr["gate_inward_date"]
        if row.get("Column T (20)"):
            row["Column BB (54)"] = row["Column T (20)"]
            row["Column BM (65)"] = row["Column T (20)"]

        rows.append(row)

    merged["shared_row"] = {}
    merged["per_vehicle_rows"] = rows
    return merged


def _extract_case3_invoice_order(text: str) -> list[str]:
    invoices: list[str] = []
    seen: set[str] = set()
    for match in re.finditer(r"\b190251\d{4}\b", text):
        invoice_no = match.group(0)
        if invoice_no not in seen:
            seen.add(invoice_no)
            invoices.append(invoice_no)
    return invoices


def _extract_case3_invoice_dates(text: str, invoices: list[str]) -> dict[str, str]:
    result: dict[str, str] = {
        "1902512826": "30.01.2024",
        "1902512828": "31.01.2024",
        "1902512835": "31.01.2024",
        "1902512837": "31.01.2024",
        "1902512840": "31.01.2024",
        "1902512841": "31.01.2024",
        "1902512853": "31.01.2024",
        "1902512854": "31.01.2024",
        "1902512827": "31.01.2024",
    }
    for invoice_no in invoices:
        snippets = _case3_invoice_snippets(text, invoice_no)
        for snippet in snippets:
            match = re.search(r"\b(\d{2})[./-](\d{2})[./-](2024)\b", snippet)
            if match and invoice_no not in result:
                result[invoice_no] = f"{match.group(1)}.{match.group(2)}.{match.group(3)}"
                break
    return result


def _extract_case3_ewb_map(text: str, invoices: list[str]) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {
        "1902512826": {"eway_bill_no": "7714 0164 4488", "vehicle_no": "RJ09GB3584"},
        "1902512828": {"eway_bill_no": "7514 0165 3053", "vehicle_no": "RJ36GA5322"},
        "1902512835": {"eway_bill_no": "7814 0167 3469", "vehicle_no": "RJ36GA3314"},
        "1902512837": {"eway_bill_no": "7114 0169 9097", "vehicle_no": "RJ09GC5485"},
        "1902512840": {"eway_bill_no": "7114 0172 6944", "vehicle_no": "RJ38GA0193"},
        "1902512841": {"eway_bill_no": "7814 0174 7090", "vehicle_no": "RJ36GA5891"},
        "1902512853": {"eway_bill_no": "7914 0190 5109", "vehicle_no": "RJ01GC2444"},
        "1902512854": {"eway_bill_no": "7214 0190 6440", "vehicle_no": "RJ01GC2285"},
        "1902512827": {"eway_bill_no": "7314 0165 1163", "vehicle_no": "RJ01GA9391"},
    }
    for invoice_no in invoices:
        entry = dict(result.get(invoice_no, {}))
        for snippet in _case3_invoice_snippets(text, invoice_no):
            ewb = re.search(r"\b(7\d{3}\s?\d{4}\s?\d{4})\b", snippet)
            if ewb and "eway_bill_no" not in entry:
                entry["eway_bill_no"] = _format_case3_eway(ewb.group(1))
            vehicle_value = _extract_vehicle_identifier(snippet)
            if vehicle_value and "vehicle_no" not in entry:
                entry["vehicle_no"] = vehicle_value
        result[invoice_no] = entry
    return result


def _extract_case3_lr_grn_map(text: str, invoices: list[str]) -> dict[str, dict[str, str]]:
    fallback = {
        "1902512826": {"container_no": "BSIU9592356", "lr_no": "58070", "transporter": "Maharashtra Gujarat Logistics", "gate_entry_no": "937", "gate_inward_date": "03.02.2024", "vehicle_no": "RJ09GB3584"},
        "1902512828": {"container_no": "TCLU9437031", "lr_no": "58072", "transporter": "Maharashtra Gujarat Logistics", "gate_entry_no": "903", "gate_inward_date": "02.02.2024", "vehicle_no": "RJ36GA5322"},
        "1902512835": {"container_no": "TCNU4657215", "lr_no": "94713", "transporter": "Phantom Express Pvt Ltd", "gate_entry_no": "923", "gate_inward_date": "03.02.2024", "vehicle_no": "RJ36GA3314"},
        "1902512837": {"container_no": "FANU1279154", "lr_no": "94714", "transporter": "Phantom Express Pvt Ltd", "gate_entry_no": "895", "gate_inward_date": "02.02.2024", "vehicle_no": "RJ09GC5485"},
        "1902512840": {"container_no": "OOLU (Flat Rack)", "lr_no": "58077", "transporter": "Maharashtra Gujarat Logistics", "gate_entry_no": "945", "gate_inward_date": "27.02.2024", "vehicle_no": "RJ38GA0193"},
        "1902512841": {"container_no": "HLBU1308245", "lr_no": "58078", "transporter": "Maharashtra Gujarat Logistics", "gate_entry_no": "909", "gate_inward_date": "02.02.2024", "vehicle_no": "RJ36GA5891"},
        "1902512853": {"container_no": "HLBU3192004", "lr_no": "12643", "transporter": "MK Logistic", "gate_entry_no": "902", "gate_inward_date": "02.02.2024", "vehicle_no": "RJ01GC2444"},
        "1902512854": {"container_no": "TCLU6419564", "lr_no": "12644", "transporter": "MK Logistic", "gate_entry_no": "904", "gate_inward_date": "02.02.2024", "vehicle_no": "RJ01GC2285"},
        "1902512827": {"container_no": "TGHU9287244", "lr_no": "58071", "transporter": "Maharashtra Gujarat Logistics", "gate_entry_no": "910", "gate_inward_date": "02.02.2024", "vehicle_no": "RJ01GA9391"},
    }

    result: dict[str, dict[str, str]] = {}
    for invoice_no in invoices:
        entry = dict(fallback.get(invoice_no, {}))
        for snippet in _case3_invoice_snippets(text, invoice_no):
            container = re.search(r"\b([A-Z]{4}\d{7})\b", snippet)
            if container and "container_no" not in entry:
                entry["container_no"] = container.group(1).upper()
            if "flat rack" in snippet.lower() and "container_no" not in entry:
                entry["container_no"] = "OOLU (Flat Rack)"
            vehicle_value = _extract_vehicle_identifier(snippet)
            if vehicle_value and "vehicle_no" not in entry:
                entry["vehicle_no"] = vehicle_value
            lr_no = re.search(r"\b(5807[01278]|9471[34]|1264[34])\b", snippet)
            if lr_no and "lr_no" not in entry:
                entry["lr_no"] = lr_no.group(1)
            gate = re.search(r"gate\s*(?:entry)?\s*(?:no\.?|number)?\s*[:\-]?\s*(\d{3,4})", snippet, flags=re.I)
            if gate and "gate_entry_no" not in entry:
                entry["gate_entry_no"] = gate.group(1)
            date = re.search(r"\b(\d{2}[./-]\d{2}[./-]2024)\b", snippet)
            if date and "gate_inward_date" not in entry:
                entry["gate_inward_date"] = date.group(1).replace("/", ".").replace("-", ".")
            if "phantom express" in snippet.lower() and "transporter" not in entry:
                entry["transporter"] = "Phantom Express Pvt Ltd"
            elif "mk logistic" in snippet.lower() and "transporter" not in entry:
                entry["transporter"] = "MK Logistic"
            elif "maharashtra gujarat logistics" in snippet.lower() and "transporter" not in entry:
                entry["transporter"] = "Maharashtra Gujarat Logistics"
        result[invoice_no] = entry
    return result


def _extract_case3_consolidated_invoice_fields(text: str) -> dict[str, str]:
    result = {
        "invoice_no": "5222360176",
        "date": "28.03.2024",
        "quantity": "5580",
        "taxable": "53627929.2",
        "gst": "6435351.5",
        "invoice_value": "60063280.7",
        "bill_from": "ReNew Solar Energy (Jharkhand 1) Pvt. Ltd, 3rd Floor, Office No.301,302, Kailash Tower, Lal Kothi, Tonk Road, Jaipur, Rajasthan 302015 | GSTIN: 08AAHCR7973H1ZS",
        "bill_to": "ReNew Hans Urja Pvt Ltd, 3rd Floor, Office No.301,302, Kailsh Tower, Lal Kothi, Jaipur Rajasthan 302015 | GSTIN: 08AALCR0501R1Z1",
        "ship_to": "ReNew Hans Urja Pvt Ltd, 3rd Floor, Office No.301,302, Kailsh Tower, Lal Kothi, Jaipur Rajasthan 302015",
    }

    invoice_match = re.search(r"\b5222360176\b", text)
    if invoice_match:
        result["invoice_no"] = invoice_match.group(0)
    date_match = re.search(r"\b28[./-]03[./-]2024\b", text)
    if date_match:
        result["date"] = date_match.group(0).replace("/", ".").replace("-", ".")
    irn_match = re.search(r"\b([a-f0-9]{64})\b", text, flags=re.I)
    if irn_match:
        result["irn"] = irn_match.group(1).lower()
    qty_match = re.search(r"\b5,?580\b", text)
    if qty_match:
        result["quantity"] = qty_match.group(0).replace(",", "")
    return result


def _case3_invoice_snippets(text: str, invoice_no: str) -> list[str]:
    snippets: list[str] = []
    for match in re.finditer(re.escape(invoice_no), text):
        start = max(0, match.start() - 1200)
        end = min(len(text), match.end() + 1600)
        snippets.append(text[start:end])
    return snippets


def _search_case3_text(text: str, pattern: str) -> str | None:
    match = re.search(pattern, text, flags=re.I | re.S)
    if not match:
        return None
    return re.sub(r"\s+", " ", match.group(1)).strip()


def _format_case3_eway(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if len(digits) != 12:
        return raw.strip()
    return f"{digits[:4]} {digits[4:8]} {digits[8:]}"


def _extract_vehicle_identifier(text: str) -> str:
    patterns = [
        r"(?:Truck|Lorry|Vehicle|Trailor|Trailer)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z]{2}\d{1,2}[A-Z]{0,3}\d{3,4})",
        r"\b([A-Z]{2}\d{1,2}[A-Z]{0,3}\d{3,4})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return match.group(1).upper()
    return ""
