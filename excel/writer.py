import re
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime


def _parse_column_key(key: str) -> int | None:
    """
    Parse a column identifier like 'Column B (2)' into a column index.
    Also handles plain column letters like 'B' or plain numbers like '2'.
    """
    # Match "Column X (N)" pattern
    match = re.match(r"Column\s+([A-Z]+)\s*\((\d+)\)", key)
    if match:
        return int(match.group(2))

    # Match plain column letter
    match = re.match(r"^([A-Z]{1,3})$", key.strip())
    if match:
        try:
            return column_index_from_string(match.group(1))
        except ValueError:
            return None

    # Match plain number
    match = re.match(r"^(\d+)$", key.strip())
    if match:
        return int(match.group(1))

    return None


def _coerce_value(value, existing_cell_value=None):
    """Convert extracted string values to appropriate Python types."""
    if value is None or value == "":
        return None

    s = str(value).strip()
    if s.upper() in ("N/A", "NA", "NONE", "NULL"):
        return "N/A"

    # Try date parsing (YYYY-MM-DD)
    date_match = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    if date_match:
        try:
            return datetime(
                int(date_match.group(1)),
                int(date_match.group(2)),
                int(date_match.group(3)),
            )
        except ValueError:
            pass

    # Try numeric
    try:
        # Remove commas
        cleaned = s.replace(",", "")
        if "." in cleaned:
            return float(cleaned)
        return int(cleaned)
    except ValueError:
        pass

    return s


def write_to_tracker(
    template_path: str,
    output_path: str,
    merged_data: dict,
    schema: dict,
    existing_data: list[dict],
    overwrite_existing: bool = False,
) -> str:
    """
    Write extracted data to the Excel tracker.

    Args:
        template_path: Path to the original Excel file
        output_path: Path to save the filled Excel file
        merged_data: {"shared_row": {...}, "per_vehicle_rows": [...]}
        schema: From excel.reader.read_tracker_schema()
        existing_data: From excel.reader.read_existing_data()
        overwrite_existing: If True, overwrite pre-filled cells

    Returns the output path.
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    shared = merged_data.get("shared_row", {})
    per_vehicle = merged_data.get("per_vehicle_rows", [])

    data_start = schema["data_start_row"]
    data_end = schema["data_end_row"]
    num_data_rows = data_end - data_start + 1

    # Write shared columns to all data rows
    for row_offset in range(num_data_rows):
        row_num = data_start + row_offset
        _write_row_data(ws, row_num, shared, overwrite_existing)

    # Write per-vehicle data by matching vehicle numbers
    if per_vehicle:
        _write_per_vehicle_data(
            ws, per_vehicle, existing_data, data_start, schema, overwrite_existing
        )

    wb.save(output_path)
    wb.close()
    return output_path


def _write_row_data(ws, row_num: int, data: dict, overwrite: bool):
    """Write a dict of column values to a specific row."""
    for key, value in data.items():
        if key.startswith("_"):
            continue
        col_idx = _parse_column_key(key)
        if col_idx is None:
            continue

        cell = ws.cell(row=row_num, column=col_idx)
        if cell.value is not None and not overwrite:
            continue  # Skip pre-filled cells

        coerced = _coerce_value(value, cell.value)
        if coerced is not None:
            cell.value = coerced


def _write_per_vehicle_data(
    ws,
    per_vehicle: list[dict],
    existing_data: list[dict],
    data_start: int,
    schema: dict,
    overwrite: bool,
):
    """Match per-vehicle data to existing rows using vehicle number."""
    # Build a lookup of vehicle_no -> row_num from existing data
    vehicle_to_row = {}
    for row_data in existing_data:
        row_num = row_data.get("_row_num")
        # Vehicle number is typically in column 63 (BK)
        vehicle_no = row_data.get(63)
        if vehicle_no and row_num:
            vehicle_to_row[str(vehicle_no).strip().upper()] = row_num

    for vehicle_data in per_vehicle:
        # Find the vehicle number in the extracted data
        vehicle_no = None
        for key, value in vehicle_data.items():
            if "vehicle" in key.lower() or key == "vehicle_no":
                vehicle_no = str(value).strip().upper()
                break

        if vehicle_no and vehicle_no in vehicle_to_row:
            row_num = vehicle_to_row[vehicle_no]
            _write_row_data(ws, row_num, vehicle_data, overwrite)
        else:
            # If no match, write to the next available row
            # Find the first row without this data
            pass  # Skip unmatched vehicles - user can review
