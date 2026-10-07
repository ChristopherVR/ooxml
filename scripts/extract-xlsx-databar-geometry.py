"""Extract native PDF vectors into a fixture; requires PyMuPDF, never renders library output."""
import argparse
import json
from pathlib import Path
import fitz

parser = argparse.ArgumentParser()
parser.add_argument('folder', type=Path)
parser.add_argument('output', type=Path)
parser.add_argument('--append', action='store_true', help='Append disjoint cases from the same Excel build')
args = parser.parse_args()
recorded = json.loads((args.folder / 'raw.json').read_text(encoding='utf-8-sig'))
for case in recorded['cases']:
    with fitz.open(args.folder / (case['id'] + '.pdf')) as document:
        page = document[0]
        drawings = page.get_drawings()
        reference = next(d['rect'] for d in drawings if d['fill'] == (1., 1., 0.))
        # A native full-length reference bar measures the usable content envelope.
        content = next(d['rect'] for d in drawings if d['fill'] == (0., 1., 0.) and reference.contains(d['rect']))
        # Native cell labels identify the five rows, including cells with no bar.
        labels = sorted((w for w in page.get_text('words') if (w[2] < reference.x0 or w[0] > reference.x1) and w[1] > 50), key=lambda w: w[1])
        assert len(labels) == 5, (case['id'], labels)
        bars = [d['rect'] for d in drawings if d['fill'] in [(0., 1., 0.), (0., 0., 1.)]]
        axes = sorted(set(round(d['rect'].x0, 3) for d in drawings if d['color'] == (1., 0., 1.)))
        assert len(axes) <= 1, (case['id'], axes)
        case['axisFraction'] = (axes[0] - content.x0) / content.width if axes else None
        case['bars'] = []
        for label in labels:
            center = (label[1] + label[3]) / 2
            bar = next((r for r in bars if abs((r.y0 + r.y1) / 2 - center) < 3), None)
            case['bars'].append(None if bar is None else {
                'start': (bar.x0 - content.x0) / content.width,
                'fraction': bar.width / content.width,
            })
recorded['evidence'] = 'Excel PDF solid-fill vectors, normalized to a full-length native reference bar; comparisons allow 0.015 for axis gaps and print quantization'
if args.append:
    existing = json.loads(args.output.read_text(encoding='utf-8'))
    assert (existing['version'], existing['build']) == (recorded['version'], recorded['build']), 'Excel build changed'
    previous = {case['id'] for case in existing['cases']}
    assert not previous.intersection(case['id'] for case in recorded['cases']), 'Duplicate native cases'
    recorded['cases'] = existing['cases'] + recorded['cases']
args.output.write_text(json.dumps(recorded, indent=2) + '\n', encoding='utf-8')
print(f"Extracted {len(recorded['cases'])} native geometry cases")
