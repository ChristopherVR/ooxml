"""Extend note bodies in the native measured independent one-row story fixtures."""
import copy,json,re,sys,zipfile
from pathlib import Path
from lxml import etree as E
out=Path(sys.argv[1]);out.mkdir(parents=True,exist_ok=True)
root=Path(__file__).resolve().parent.parent
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+ns['w']+'}'
cases=[(profile+'-'+str(columns),source,lengths,columns)for profile,source,lengths in [('both','two-inner',[40,40]),('consumed','consumed-inner',[40,40]),('enclosing','consumed-enclosing',[40,40]),('three','consumed-three',[40,3,40])]for columns in [2,3]]
cases += [('enclosing-2-sixty','consumed-enclosing',[60,60],2),('enclosing-3-sixty','consumed-enclosing',[60,60],3),('three-3-sixty','consumed-three',[60,3,60],3)]
records=[]
for name,source,lengths,columns in cases:
 name='column-note-'+name
 with zipfile.ZipFile(root/'src/core/docx/__fixtures__/multiple-inner-notes'/('multi-note-'+source+'.docx'))as z:parts={p:z.read(p)for p in z.namelist()}
 notes=E.fromstring(parts['word/footnotes.xml'])
 for number,note in enumerate([n for n in notes if int(n.get(W+'id','-1'))>0]):
  paragraphs=note.findall('w:p',ns);last=paragraphs[-1]
  for line in range(len(paragraphs),lengths[number]):
   p=copy.deepcopy(last)
   for t in p.findall('.//w:t',ns):t.text=re.sub(r'N\d+L\d+',f'N{number}L{line:02}',t.text or '')
   note.append(p)
 parts['word/footnotes.xml']=E.tostring(notes,xml_declaration=True,encoding='UTF-8',standalone=True)
 with zipfile.ZipFile(out/(name+'-input.docx'),'w',zipfile.ZIP_DEFLATED)as z:
  for p,data in parts.items():z.writestr(p,data)
 records.append({'name':name,'source':source,'noteLines':lengths,'noteColumns':columns})
(out/'input-cases.json').write_text(json.dumps(records,indent=2)+'\n',encoding='utf-8')
