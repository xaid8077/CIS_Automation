# tools/embed_templates.py

import base64
from pathlib import Path

TEMPLATES = {
    "IL": "templates/IL.xlsx",
    "IO": "templates/IO.xlsx",
}

output_lines = [
    "# AUTO-GENERATED FILE — DO NOT EDIT MANUALLY\n",
]

for name, path in TEMPLATES.items():
    data = Path(path).read_bytes()
    b64 = base64.b64encode(data).decode()

    output_lines.append(f"{name}_TEMPLATE_B64 = '''{b64}'''\n\n")

Path("utils/embedded_templates.py").write_text("".join(output_lines))

print("✅ Embedded templates written to utils/embedded_templates.py")