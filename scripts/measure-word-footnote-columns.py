"""Read native Word PDFs from record-word-footnote-columns.ps1 using PyMuPDF."""
import json
import re
import sys
from pathlib import Path
import fitz

root = Path(sys.argv[1])
records = json.loads((root / 'native-cases.json').read_text(encoding='utf-8-sig'))
for case in records:
    case['pages'] = []
    with fitz.open(root / (case['name'] + '.pdf')) as document:
        for page in document:
            columns = []
            captions = []
            body_lines = []
            for block in page.get_text('dict')['blocks']:
                for line in block.get('lines', []):
                    text = ''.join(span['text'] for span in line['spans']).strip()
                    match = re.search(r'N\dL\d\d', text)
                    if 'body' in text:
                        body_lines.append({'text': text, 'leftPt': round(line['bbox'][0], 3), 'baselinePt': round(line['spans'][-1]['origin'][1], 3)})
                    if not match:
                        if 'Sources' in text or 'Continued' in text:
                            captions.append(text)
                        continue
                    left = line['bbox'][0]
                    column = next((column for column in columns if abs(column['leftPt'] - left) < 0.5), None)
                    if column is None:
                        column = {'leftPt': round(left, 3), 'lines': []}
                        columns.append(column)
                    column['lines'].append({'id': match[0], 'baselinePt': round(line['spans'][-1]['origin'][1], 3)})
            rules = [list(round(value, 3) for value in drawing['rect']) for drawing in page.get_drawings()
                     if drawing['rect'].height < 2 and drawing['rect'].width > 20]
            case['pages'].append({'widthPt': page.rect.width, 'heightPt': page.rect.height,
                                  'columns': sorted(columns, key=lambda column: column['leftPt']),
                                  'separatorCaptions': captions, 'rulesPt': rules, 'bodyLines': body_lines})
output = {'provenance': 'Owned Word.Application COM 16.0; synthetic documents; native PDF coordinates in points; generator scripts/record-word-footnote-columns.ps1',
          'cases': records}
(root / 'measured-cases.json').write_text(json.dumps(output, indent=2) + '\n', encoding='utf-8')
print(root / 'measured-cases.json')
