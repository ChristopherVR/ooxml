"""Add native cell-line/reference measurements to the shared footnote PDF measurements."""
import json
import re
import subprocess
import sys
from pathlib import Path
import fitz
root = Path(sys.argv[1])
subprocess.run([sys.executable, str(Path(__file__).with_name('measure-word-footnote-columns.py')), str(root)], check=True)
path = root / 'measured-cases.json'
measured = json.loads(path.read_text(encoding='utf-8'))
for case in measured['cases']:
    with fitz.open(root / (case['name'] + '.pdf')) as document:
        for page, record in zip(document, case['pages']):
            cells = []
            for block in page.get_text('dict')['blocks']:
                for line in block.get('lines', []):
                    text = ''.join(span['text'] for span in line['spans'])
                    match = re.search(r'T\dR\d+C\d+L\d+', text)
                    if match:
                        cells.append({'id': match[0], 'leftPt': round(line['bbox'][0], 3),
                                      'baselinePt': round(line['spans'][-1]['origin'][1], 3)})
            record['cellLines'] = cells
path.write_text(json.dumps(measured, indent=2) + '\n', encoding='utf-8')
print(path)
