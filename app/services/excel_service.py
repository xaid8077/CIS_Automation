# app/services/excel_service.py

from io import BytesIO
from datetime import datetime

from app.utils.excel_writer import write_workbook, write_io_workbook


def generate_excel(payload, doc_type):
    output = BytesIO()

    if doc_type == "Instrument List":
        write_workbook(payload, output)
        prefix = "Instrument_List"

    elif doc_type == "IO List":
        write_io_workbook(payload, output)
        prefix = "IO_List"

    else:
        raise ValueError("Unsupported document type")

    output.seek(0)
    return output, prefix


def build_filename(project, location, prefix, revision):
    loc_tag = ""

    if location:
        loc_tag = (
            location.code.strip()
            if location.code
            else location.name[:12].replace(" ", "_")
        )

    return (
        f"{project.display_name.replace(' ', '_')}"
        f"_{loc_tag}_{prefix}_Rev{revision}"
        f"_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )