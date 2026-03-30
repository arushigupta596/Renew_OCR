import re


def validate_cross_references(all_extracted: dict[str, dict]) -> list[dict]:
    """
    Validate that common fields match across documents.

    Returns a list of validation results:
    [{"field": str, "status": "pass"|"warning"|"fail", "message": str}, ...]
    """
    results = []

    # Collect all values by column key across documents
    all_values = {}
    for filename, data in all_extracted.items():
        if isinstance(data, dict) and "_error" not in data and "_parse_error" not in data:
            # Handle both flat and shared/per_row formats
            flat = data.get("shared", data) if "shared" in data else data
            for key, value in flat.items():
                if key.startswith("_"):
                    continue
                if key not in all_values:
                    all_values[key] = []
                all_values[key].append({"filename": filename, "value": value})

    # Check columns that appear in multiple documents
    for key, entries in all_values.items():
        if len(entries) > 1:
            values = [_normalize_value(e["value"]) for e in entries]
            unique = set(values)
            if len(unique) == 1:
                results.append({
                    "field": key,
                    "status": "pass",
                    "message": f"Consistent across {len(entries)} documents: {values[0]}",
                })
            else:
                sources = ", ".join(
                    f"{e['filename']}: {e['value']}" for e in entries
                )
                results.append({
                    "field": key,
                    "status": "warning",
                    "message": f"Mismatch across documents - {sources}",
                })

    # Check for common cross-reference patterns
    _check_linkage_fields(all_extracted, results)

    return results


def _normalize_value(val) -> str:
    """Normalize a value for comparison."""
    s = str(val).strip().upper()
    # Remove extra whitespace
    s = re.sub(r"\s+", " ", s)
    return s


def _check_linkage_fields(all_extracted: dict[str, dict], results: list[dict]):
    """Check specific linkage patterns (invoice numbers, BL numbers, etc.)."""
    # Collect all unique numeric-looking values and string identifiers
    invoice_numbers = set()
    bl_numbers = set()
    boe_numbers = set()

    for filename, data in all_extracted.items():
        if isinstance(data, dict) and "_error" not in data:
            flat = data.get("shared", data) if "shared" in data else data
            for key, value in flat.items():
                key_lower = key.lower()
                val_str = str(value).strip()
                if "invoice" in key_lower and "no" in key_lower:
                    invoice_numbers.add(val_str)
                elif "bl" in key_lower and ("number" in key_lower or "no" in key_lower):
                    bl_numbers.add(val_str)
                elif "boe" in key_lower and ("number" in key_lower or "no" in key_lower):
                    boe_numbers.add(val_str)

    if len(invoice_numbers) > 1:
        results.append({
            "field": "Invoice Numbers",
            "status": "warning",
            "message": f"Multiple invoice numbers found: {invoice_numbers}",
        })
    elif len(invoice_numbers) == 1:
        results.append({
            "field": "Invoice Numbers",
            "status": "pass",
            "message": f"Consistent invoice number: {invoice_numbers.pop()}",
        })

    if len(bl_numbers) > 1:
        results.append({
            "field": "BL Numbers",
            "status": "warning",
            "message": f"Multiple BL numbers found: {bl_numbers}",
        })

    if len(boe_numbers) > 1:
        results.append({
            "field": "BoE Numbers",
            "status": "warning",
            "message": f"Multiple BoE numbers found: {boe_numbers}",
        })
