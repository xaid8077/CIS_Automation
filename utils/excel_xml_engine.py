# utils/excel_xml_engine.py
"""
DEPRECATED — DO NOT USE.

This module has been superseded by utils/excel_writer.py which provides
a cleaner, fully self-contained XML engine with proper shared-string
management, date serialisation, and cell-creation for missing rows/cells.

Safe to delete once you have confirmed nothing in your deployment imports
this file directly.

All public functions below raise NotImplementedError to catch any stale
import at startup rather than silently producing corrupt output.
"""


def patch_excel(*args, **kwargs):
    raise NotImplementedError(
        "excel_xml_engine.patch_excel() is deprecated. "
        "Use utils.excel_writer._process() via excel_service.generate() instead."
    )


def map_sheet_names(*args, **kwargs):
    raise NotImplementedError(
        "excel_xml_engine.map_sheet_names() is deprecated. "
        "Use utils.excel_writer._map_sheets() instead."
    )