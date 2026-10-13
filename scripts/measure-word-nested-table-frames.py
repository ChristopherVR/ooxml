"""Measure native nested cell fills, border segments and reference text in PDF points."""
import json,re,sys
from pathlib import Path
import fitz
root=Path(sys.argv[1])
cases=json.loads((root/'native-cases.json').read_text(encoding='utf-8-sig'))
for case in cases:
 case['measurements']=[]
 with fitz.open(root/(case['name']+'.pdf')) as pdf:
  for page in pdf:
   fills=[];borders=[];lines=[]
   for drawing in page.get_drawings():
    color=drawing['color'];fill=drawing['fill']
    if fill and tuple(round(c*255) for c in fill) in [(255,255,0),(0,255,255),(0,255,0),(255,0,255)]:
     fills.append({'color':'#'+''.join(f'{round(c*255):02X}' for c in fill),'rectPt':[round(v,3) for v in drawing['rect']]})
    if fill and fill[0]>.99 and fill[1]<.01 and fill[2]<.01:
     borders.append({'rectPt':[round(v,3) for v in drawing['rect']]})
    if color and color[0]>.99 and color[1]<.01 and color[2]<.01:
     for item in drawing['items']:
      if item[0]=='l':borders.append({'fromPt':[round(v,3) for v in item[1]],'toPt':[round(v,3) for v in item[2]],'widthPt':round(drawing['width'],3)})
   for block in page.get_text('dict')['blocks']:
    for line in block.get('lines',[]):
     text=''.join(span['text'] for span in line['spans'])
     if re.search(r'T\dR\d+C\d+L[A-Z]?\d+|N0L\d+|D[01]',text):lines.append({'text':text,'leftPt':round(line['bbox'][0],3),'baselinePt':round(line['spans'][-1]['origin'][1],3)})
   case['measurements'].append({'fills':fills,'borders':borders,'lines':lines})
(root/'native-measurements.json').write_text(json.dumps({'unit':'PDF points','cases':cases},indent=2)+'\n',encoding='utf-8')
for case in cases:print(case['name'],[(p['fills'],len(p['borders']),[l for l in p['lines'] if 'T1R0C0L00' in l['text']])for p in case['measurements']])
