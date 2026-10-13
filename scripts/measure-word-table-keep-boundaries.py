import json,hashlib,pathlib,sys
root=pathlib.Path(sys.argv[1]);cases=json.loads((root/'probe-cases.json').read_text());native=json.loads((root/'native-measurements.json').read_text());out=root
records=[]
for c in cases:
 n=next(n for n in native['cases'] if n['name']==c['name'])
 pages=[]
 for p in n['measurements']:
  lines=[l for l in p['lines'] if l['text'].startswith('T1')]
  borders=[b['rectPt'] for b in p['borders'] if 'rectPt'in b]
  pages.append({'body':[l['text'].strip() for l in lines],'topPt':min((b[1] for b in borders),default=None),'bottomPt':max((b[3] for b in borders),default=None),'firstBaselinePt':lines[0]['baselinePt'] if lines else None})
 records.append({**c,'pages':pages,'sourceSHA256':hashlib.sha256((root/(c['name']+'.docx')).read_bytes()).hexdigest(),'inputSHA256':hashlib.sha256((root/(c['name']+'-input.docx')).read_bytes()).hexdigest()})
(out/'native-boundaries.json').write_text(json.dumps({'unit':'PDF points','wordVersion':native['cases'][0]['wordVersion'],'wordBuild':native['cases'][0]['wordBuild'],'cases':records},indent=2)+'\n',encoding='utf8')
print('compact cases',len(records))