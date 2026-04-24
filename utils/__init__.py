# utils/__init__.py
"""
utils/
──────
Utility package.

Modules
───────
excel_writer.py       — XML-level Excel generation (IL + IO workbooks).
embedded_templates.py — Auto-generated base-64 blobs (run xlto64.py to
                        regenerate after editing an .xlsx template).
validator.py          — Business-rule validation (cross-row, tag checks).
rbac.py               — RBAC decorators for web and API routes.

Removed
───────
excel_xml_engine.py   — Superseded by excel_writer.py.  Delete this file;
                        nothing in the current codebase imports it.
parser.py             — Superseded; payload parsing moved into
                        schemas/payload.py (Marshmallow).
"""