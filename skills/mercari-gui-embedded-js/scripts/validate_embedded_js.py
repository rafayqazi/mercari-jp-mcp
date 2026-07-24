#!/usr/bin/env python3
"""Extract runtime JS from mercari_gui.HTML_PAGE and verify syntax with Node."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(ROOT))

import mercari_gui  # noqa: E402

html = mercari_gui.HTML_PAGE
start = html.find("<script>") + len("<script>")
end = html.find("</script>")
if start < len("<script>") or end == -1:
    print("ERROR: <script> block not found in HTML_PAGE")
    sys.exit(1)

js = html[start:end]
out = ROOT / "_runtime_js_validate.tmp.js"
out.write_text(js, encoding="utf-8")

try:
    subprocess.run(
        ["node", "--check", str(out)],
        check=True,
        cwd=str(ROOT),
    )
except FileNotFoundError:
    print("WARN: node not found; wrote", out)
    sys.exit(0)
except subprocess.CalledProcessError:
    print("ERROR: JavaScript syntax check failed — see", out)
    sys.exit(1)
finally:
    if out.exists():
        out.unlink()

print("OK: embedded JavaScript syntax valid")
