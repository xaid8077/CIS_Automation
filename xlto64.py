# xlto64.py
"""
xlto64.py
─────────
Embed Excel templates as base-64 strings inside utils/embedded_templates.py.

Usage (run once after changing a template file):
    python xlto64.py

The generated file is imported by utils/excel_writer.py at runtime so the
application never needs the raw .xlsx files on the production filesystem.
"""

import base64
from pathlib import Path

# Map variable name → source template path
TEMPLATES = {
    "IL": "templates/IL.xlsx",
    "IO": "templates/IO.xlsx",
}

output_lines = [
    "# AUTO-GENERATED FILE — DO NOT EDIT MANUALLY\n",
    "# Regenerate with:  python xlto64.py\n",
    "\n",
]

for name, path in TEMPLATES.items():
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(
            f"Template not found: {path}\n"
            f"Make sure both templates/IL.xlsx and templates/IO.xlsx exist "
            f"before running this script."
        )
    b64 = base64.b64encode(src.read_bytes()).decode()
    output_lines.append(f"{name}_TEMPLATE_B64 = '{b64}'\n\n")

out_path = Path("utils/embedded_templates.py")
out_path.parent.mkdir(parents=True, exist_ok=True)
out_path.write_text("".join(output_lines))

print(f"✅  Embedded templates written to {out_path}")
print(f"    IL: {Path(TEMPLATES['IL']).stat().st_size:,} bytes")
print(f"    IO: {Path(TEMPLATES['IO']).stat().st_size:,} bytes")