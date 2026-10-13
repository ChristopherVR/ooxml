"""Record native outer/nested border rectangles, fills and cell ink in PDF points."""
import json,re,sys
from pathlib import Path
import fitz
root=Path(sys.argv[1]);cases=json.loads((root/'native-cases.json').read_text(encoding='utf-8-sig'))
for case in cases:
 case['measurements']=[]
 with fitz.open(root/(case['name']+'.pdf')) as pdf:
  for page in pdf:
   borders=[];fills=[];lines=[]
   for drawing in page.get_drawings():
    fill=drawing['fill'];rgb=tuple(round(c*255)for c in fill)if fill else None
    if rgb in [(255,0,0),(0,0,255)]:borders.append({'role':'outer'if rgb==(255,0,0)else'nested','rectPt':[round(v,3)for v in drawing['rect']]})
    if rgb in [(255,255,0),(0,255,255),(0,255,0),(255,0,255)]:fills.append({'color':'#'+''.join(f'{v:02X}'for v in rgb),'rectPt':[round(v,3)for v in drawing['rect']]})
   for block in page.get_text('dict')['blocks']:
    for line in block.get('lines',[]):
     text=''.join(span['text']for span in line['spans'])
     if re.search(r'T\dR\d+C\d+L\d+|N0L\d+|Outer cell',text):lines.append({'text':text,'leftPt':round(line['bbox'][0],3),'baselinePt':round(line['spans'][-1]['origin'][1],3)})
   measurement={'borders':borders,'fills':fills,'lines':lines}
   if case['name'].startswith('enclosing-border-') and case['name']!='enclosing-border-split':
    # Colored frame/fill pixels are excluded: this crop measures retained dark glyph ink.
    crop=fitz.Rect(44,58,103,69);scale=4
    pix=page.get_pixmap(matrix=fitz.Matrix(scale,scale),clip=crop,alpha=False)
    samples=pix.samples;ys=[]
    for y in range(pix.height):
     for x in range(pix.width):
      offset=(y*pix.width+x)*pix.n
      if max(samples[offset:offset+3])<80:ys.append(crop.y0+y/scale)
    measurement['trailingGlyphRaster']={'cropPt':list(crop),'scale':scale,'darkPixels':len(ys),'minYPt':min(ys)if ys else None,'maxYPt':max(ys)if ys else None}
   case['measurements'].append(measurement)
(root/'native-measurements.json').write_text(json.dumps({'unit':'PDF points','cases':cases},indent=2)+'\n',encoding='utf-8')
for case in cases:print(case['name'],case['pages'])
