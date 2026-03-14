from pathlib import Path
from typing import Any, Dict, Iterable, List

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
HEADER_FONT = Font(bold=True)
CENTER = Alignment(vertical="top", wrap_text=True)


def _write_sheet(ws, rows: List[Dict[str, Any]], title: str) -> None:
    ws.title = title

    if not rows:
        ws.append(["No data"])
        return

    headers = list(rows[0].keys())
    ws.append(headers)

    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER

    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = CENTER

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 14), 40)


def _flatten_signals(
    equipment_rows: Iterable[Dict[str, Any]],
    section_name: str,
) -> List[Dict[str, Any]]:
    flat_rows: List[Dict[str, Any]] = []

    for eq in equipment_rows:
        for signal in eq.get("signals", []):
            flat_rows.append(
                {
                    "section": section_name,
                    "tag_no": eq.get("tag_no", ""),
                    "instrument": eq.get("instrument", ""),
                    "service_description": eq.get("service_description", ""),
                    "signal_description": signal.get("signal_description", ""),
                    "io_type": signal.get("io_type", ""),
                }
            )

    return flat_rows


def write_workbook(payload: Dict[str, Any], output_path: Path) -> None:
    wb = Workbook()

    # Sheet 1: Header
    ws_header = wb.active
    ws_header.title = "Header"
    header_rows = [{"field": k, "value": v} for k, v in payload.get("header", {}).items()]
    _write_sheet(ws_header, header_rows, "Header")

    # Sheet 2: Field Instruments
    ws_fi = wb.create_sheet("Field Instruments")
    _write_sheet(ws_fi, payload.get("field_instruments", []), "Field Instruments")

    # Sheet 3: Electrical Equipment
    electrical_parents = [
        {
            "tag_no": row.get("tag_no", ""),
            "instrument": row.get("instrument", ""),
            "service_description": row.get("service_description", ""),
        }
        for row in payload.get("electrical_equipment", [])
    ]
    ws_el_parent = wb.create_sheet("Electrical Equip")
    _write_sheet(ws_el_parent, electrical_parents, "Electrical Equip")

    # Sheet 4: Electrical Signals
    electrical_signals = _flatten_signals(payload.get("electrical_equipment", []), "Electrical")
    ws_el_signals = wb.create_sheet("Electrical Signals")
    _write_sheet(ws_el_signals, electrical_signals, "Electrical Signals")

    # Sheet 5: MOV Equipment
    mov_parents = [
        {
            "tag_no": row.get("tag_no", ""),
            "instrument": row.get("instrument", ""),
            "service_description": row.get("service_description", ""),
        }
        for row in payload.get("mov_equipment", [])
    ]
    ws_mov_parent = wb.create_sheet("MOV Equip")
    _write_sheet(ws_mov_parent, mov_parents, "MOV Equip")

    # Sheet 6: MOV Signals
    mov_signals = _flatten_signals(payload.get("mov_equipment", []), "MOV")
    ws_mov_signals = wb.create_sheet("MOV Signals")
    _write_sheet(ws_mov_signals, mov_signals, "MOV Signals")

    # Sheet 7: Cable Schedule
    ws_cable = wb.create_sheet("Cable Schedule")
    _write_sheet(ws_cable, payload.get("cable_schedule", []), "Cable Schedule")

    wb.save(output_path)