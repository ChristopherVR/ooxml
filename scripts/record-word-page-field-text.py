"""Extract visible native Word page-field markers from its exported PDFs (requires pypdf)."""
import json
import re
import sys
from pathlib import Path

from pypdf import PdfReader

directory = Path(sys.argv[1]).resolve()
evidence = json.loads((directory / "application.json").read_text(encoding="utf-8-sig"))
evidence["measurement"] = "visible text in Word ExportAsFixedFormat PDFs, extracted with pypdf"
evidence["cases"] = []
for source in sorted(directory.glob("*.pdf")):
    pages = []
    for page in PdfReader(source).pages:
        text = " ".join(page.extract_text().split())
        header = re.search(r"H-[AB]-[a-z-]+", text)
        footer = re.search(r"F-[AB]-[a-z-]+ P:[0-9ivxlcdm]+ S:[0-9]+", text)
        if header is None or footer is None:
            raise ValueError(f"Missing page markers in {source.name}")
        pages.append({"header": header.group(), "footer": footer.group()})
    evidence["cases"].append({"name": source.stem, "pages": pages})
(directory / "evidence.json").write_text(json.dumps(evidence, indent=2) + "\n", encoding="utf-8")
print(json.dumps(evidence, indent=2))
