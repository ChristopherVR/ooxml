"""Create synthetic interacting references from the recorded one-row nested source."""
import copy,json,re,sys,zipfile
from pathlib import Path
from lxml import etree as E
out=Path(sys.argv[1]);out.mkdir(parents=True,exist_ok=True)
root=Path(__file__).resolve().parent.parent
source=root/'src/core/docx/__fixtures__/ordinary-table-borders/enclosing-border-split.docx'
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+ns['w']+'}'
cases=[('three-inner',[(0,0),(0,15),(1,16)],20),('consumed-three-inner',[(0,0),(0,28),(1,29)],30),('two-inner',[(0,0),(1,15)],20),('two-reversed',[(1,0),(0,15)],20),('two-same-story',[(0,0),(0,15)],20),('two-enclosing',[(0,0),(2,15)],20),('three-stories',[(0,0),(1,15),(2,16)],20),('consumed-inner',[(0,0),(1,28)],30),('consumed-enclosing',[(0,0),(2,28)],30),('consumed-three',[(0,0),(1,28),(2,29)],30)]
records=[]
for name,markers,count in cases:
 name='multi-note-'+name
 with zipfile.ZipFile(source)as z:parts={p:z.read(p)for p in z.namelist()}
 doc=E.fromstring(parts['word/document.xml']);notes=E.fromstring(parts['word/footnotes.xml'])
 outer=doc.find('w:body/w:tbl/w:tr',ns)
 inner=outer.find('w:tc/w:tbl/w:tr',ns)
 paragraphs=[cell.find('w:p',ns)for cell in inner.findall('w:tc',ns)]+[outer.findall('w:tc',ns)[1].find('w:p',ns)]
 ref=copy.deepcopy(doc.find('.//w:r[w:footnoteReference]',ns))
 for run in doc.findall('.//w:r[w:footnoteReference]',ns):run.getparent().remove(run)
 for index,p in enumerate(paragraphs):
  runs=p.findall('w:r[w:t]',ns);last=runs[-1];existing=len(runs)
  for line in range(existing,count):
   run=copy.deepcopy(last)
   for br in run.findall('w:br',ns):run.remove(br)
   br=E.Element(W+'br');run.insert(list(run).index(run.find('w:t',ns)),br)
   run.find('w:t',ns).text=f'T{1 if index<2 else 0}R0C{index if index<2 else 1}L{line:02}'
   p.append(run)
 template=copy.deepcopy(next(n for n in notes if int(n.get(W+'id','-1'))>0))
 for n in list(notes):
  if int(n.get(W+'id','-1'))>0:notes.remove(n)
 for number,(cell,line)in enumerate(markers,1):
  p=paragraphs[cell];target=next(run for run in p.findall('w:r[w:t]',ns)if run.find('w:t',ns).text.strip().endswith(f'L{line:02}'))
  run=copy.deepcopy(ref);run.find('w:footnoteReference',ns).set(W+'id',str(number));p.insert(list(p).index(target)+1,run)
  note=copy.deepcopy(template);note.set(W+'id',str(number))
  for text in note.findall('.//w:t',ns):text.text=re.sub(r'N0L',f'N{number-1}L',text.text or '')
  notes.append(note)
 parts['word/document.xml']=E.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)
 parts['word/footnotes.xml']=E.tostring(notes,xml_declaration=True,encoding='UTF-8',standalone=True)
 with zipfile.ZipFile(out/(name+'-input.docx'),'w',zipfile.ZIP_DEFLATED)as z:
  for p,data in parts.items():z.writestr(p,data)
 records.append({'name':name,'markers':markers,'storyLines':count})
(out/'input-cases.json').write_text(json.dumps(records,indent=2)+'\n',encoding='utf-8')
