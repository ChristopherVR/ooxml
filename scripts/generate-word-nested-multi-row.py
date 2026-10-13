"""Build structurally bounded nested-table probes from the existing native enclosing fixture."""
import copy, json, sys
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
ROOT = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
CASES = [('short', ((3, 2), (3, 2))), ('row-boundary', ((16, 14), (3, 2))), ('first-row-split', ((20, 18), (6, 4))), ('second-row-split', ((5, 3), (20, 18))), ('both-long', ((20, 18), (20, 16)))]
profile = sys.argv[2] if len(sys.argv) > 2 else 'two'
if profile not in ['two', 'three', 'structural']: raise ValueError('Profile must be two, three or structural')
if profile == 'three':
 CASES = [('short', ((3, 2), (3, 2), (3, 2))), ('row-boundary', ((16, 14), (3, 2), (3, 2))), ('first-row-split', ((20, 18), (4, 2), (4, 2))), ('middle-row-split', ((4, 2), (20, 18), (4, 2))), ('last-row-split', ((4, 2), (4, 2), (20, 18))), ('all-long', ((20, 18), (20, 16), (20, 18)))]
prefix = 'nested-three' if profile == 'three' else 'nested-multi'
if profile == 'structural':
 CASES = []
 for count in [4, 8]:
  for name in ['short', 'boundary', 'first', 'middle', 'last', 'all-long', 'minimal', 'empty-first', 'empty-middle', 'empty-last', 'all-empty']:
   rows = [(3, 2)] * count
   if name == 'boundary': rows[0] = (16, 14)
   if name in ['first', 'middle', 'last']:
    row = 0 if name == 'first' else count // 2 if name == 'middle' else count - 1
    rows[row] = (20, 18)
   if name == 'all-long': rows = [(20, 18 if i % 2 == 0 else 16) for i in range(count)]
   if name in ['minimal', 'empty-first', 'empty-middle', 'empty-last', 'all-empty']:
    rows = [(1, 1)] * count
    if name.startswith('empty-'):
     row = 0 if name == 'empty-first' else count // 2 if name == 'empty-middle' else count - 1
     rows[row] = (0, 0)
    if name == 'all-empty': rows = [(0, 0)] * count
   CASES.append((f'{count}-{name}', tuple(rows)))
 prefix = 'nested-structural'
source = ROOT / 'src/core/docx/__fixtures__/ordinary-table-borders/enclosing-border-split.docx'
with ZipFile(source) as z: parts = {n: z.read(n) for n in z.namelist()}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; NS = {'w': W}
records = []
for name, counts in CASES:
 x = E.fromstring(parts['word/document.xml'])
 inner = x.xpath('//w:body/w:tbl/w:tr/w:tc/w:tbl', namespaces=NS)[0]
 for border in x.xpath('//w:tblPr/w:tblBorders/*', namespaces=NS):
  border.set('{'+W+'}color', 'FF0000' if border.getparent().getparent().getparent() is inner else '000000')
 template = copy.deepcopy(inner.find('w:tr', NS))
 for row in list(inner.findall('w:tr', NS)): inner.remove(row)
 for ri, cells in enumerate(counts):
  row = copy.deepcopy(template)
  for trpr in row.findall('w:trPr', NS): row.remove(trpr)
  for ci, cell in enumerate(row.findall('w:tc', NS)):
   pr = cell.find('w:tcPr', NS)
   for tag in ['vMerge', 'vAlign']:
    for old in pr.findall('w:'+tag, NS): pr.remove(old)
   paragraph = cell.find('w:p', NS)
   ppr = copy.deepcopy(paragraph.find('w:pPr', NS))
   rpr = copy.deepcopy(paragraph.find('w:r/w:rPr', NS))
   for old in list(cell):
    if old.tag != '{'+W+'}tcPr': cell.remove(old)
   paragraph = E.SubElement(cell, '{'+W+'}p'); paragraph.append(ppr)
   for line in range(cells[ci]):
    run = E.SubElement(paragraph, '{'+W+'}r'); run.append(copy.deepcopy(rpr))
    if line: E.SubElement(run, '{'+W+'}br')
    E.SubElement(run, '{'+W+'}t').text = f'T1R{ri}C{ci}L{line:02}'
  inner.append(row)
 for ref in x.xpath('//w:footnoteReference|//w:lastRenderedPageBreak', namespaces=NS): ref.getparent().remove(ref)
 for text in x.xpath('//w:t[text()="Outer cell"]', namespaces=NS): text.text = 'T2R0C0L00'
 data = dict(parts); data['word/document.xml'] = E.tostring(x, encoding='utf-8', xml_declaration=True, standalone=True)
 path = OUT / (prefix+'-'+name+'-input.docx')
 with ZipFile(path, 'w', ZIP_DEFLATED) as z:
  for part, content in data.items(): z.writestr(part, content)
 records.append({'name': prefix+'-'+name, 'lineCounts': counts})
(OUT/'probe-cases.json').write_text(json.dumps(records, indent=2)+'\n', encoding='utf-8')
