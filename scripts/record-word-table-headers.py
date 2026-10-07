"""Record repeated header counts from native Word PDFs, rather than source ranges."""
import json
import re
import sys
from pathlib import Path

from pypdf import PdfReader

directory = Path(sys.argv[1]).resolve()
source = directory / "evidence.json"
evidence = json.loads(source.read_text(encoding="utf-8-sig"))
for case in evidence["cases"]:
    if not case["name"].startswith("table-header"):
        continue
    pages = PdfReader(directory / (case["name"] + ".pdf")).pages
    assert len(pages) == case["pages"]
    case["headerCounts"] = [len(re.findall(r"\bRow1\b", page.extract_text())) for page in pages]
    print(case["name"], case["headerCounts"])
source.write_text(json.dumps(evidence, indent=2) + "\n", encoding="utf-8")
