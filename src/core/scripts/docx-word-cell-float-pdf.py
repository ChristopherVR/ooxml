# Reads the PDFs `docx-word-cell-float-probe.ps1` exported from Word and records, in points from
# each page's top-left: the red object's rectangle (a shape) or the green chart area's (a chart),
# the table's horizontal rules (its row edges) and the words the tests read. Needs PyMuPDF.
#
#   python src/core/scripts/docx-word-cell-float-pdf.py <pdf dir> <word-positions.json>
import json
import pathlib
import sys

import fitz  # PyMuPDF

source = pathlib.Path(sys.argv[1])
target = pathlib.Path(sys.argv[2])
build = (source / 'word-build.txt').read_text(encoding='utf-8-sig').strip()
r2 = lambda value: round(value, 2)
cases = {}
for pdf in sorted(source.glob('*.pdf')):
    document = fitz.open(pdf)
    page = document[0]
    objects = []
    rules = set()
    for drawing in page.get_drawings():
        fill = drawing.get('fill')
        rect = drawing['rect']
        if fill and tuple(round(c, 2) for c in fill) in ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0)):
            objects.append({'x': r2(rect.x0), 'y': r2(rect.y0), 'width': r2(rect.width), 'height': r2(rect.height)})
            continue
        # Table borders: thin horizontal strokes or fills at least a cell wide.
        if rect.width > 100 and rect.height < 2:
            rules.add(r2((rect.y0 + rect.y1) / 2))
    # Words as [text, x, y, end x]: every word for the wrap and chart cases, else the anchor's line.
    words = [
        [w[4], r2(w[0]), r2(w[1]), r2(w[2])]
        for w in page.get_text('words')
        if pdf.stem.startswith(('wrap-', 'chart-')) or w[4] in ('Anchor', 'after', 'tail', 'anchor', 'Body')
    ]
    cases[pdf.stem] = {
        'pages': len(document),
        'objects': objects,
        'rules': sorted(rules),
        'words': words,
    }
lines = ',\n'.join(f'  {json.dumps(name)}: {json.dumps(case, separators=(",", ":"))}' for name, case in cases.items())
header = f'{{\n "wordBuild": {json.dumps(build)},\n "unit": "points from the page top-left",\n "cases": {{\n'
target.write_text(f'{header}{lines}\n }}\n}}\n', encoding='utf-8')
print(f'{len(cases)} cases -> {target}')
