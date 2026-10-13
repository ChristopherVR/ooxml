import copy,json,sys
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
ROOT=Path(__file__).resolve().parent.parent;OUT=Path(sys.argv[1]);OUT.mkdir(parents=True,exist_ok=True)
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda n:'{'+W+'}'+n
with ZipFile(ROOT/'src/core/docx/__fixtures__/ordinary-table-borders/enclosing-border-split.docx') as z:parts={n:z.read(n) for n in z.namelist()}
cases=[]
def paragraph(label,count,keep=False):
 p=E.Element(q('p'));pr=E.SubElement(p,q('pPr'))
 E.SubElement(pr,q('spacing'),{q('before'):'0',q('after'):'0',q('line'):'240',q('lineRule'):'exact'})
 E.SubElement(pr,q('widowControl'),{q('val'):'0'})
 E.SubElement(pr,q('keepLines'),{q('val'):'1' if keep else '0'})
 for i in range(count):
  r=E.SubElement(p,q('r'));rp=E.SubElement(r,q('rPr'));E.SubElement(rp,q('rFonts'),{q('ascii'):'Times New Roman',q('hAnsi'):'Times New Roman'});E.SubElement(rp,q('sz'),{q('val'):'20'})
  if i:E.SubElement(r,q('br'))
  E.SubElement(r,q('t')).text=f'T1R0C0L{label}{i:02}'
 return p
for profile,body,prefix,target,suffix in [('partial',0,15,6,2),('oversize',0,3,25,2),('consumed',0,30,6,2),('fresh',15,0,6,2),('oversize-fresh',0,0,25,2),('oversize-consumed',0,30,25,2),('oversize-prefilled',15,0,25,2),('mirror-partial',0,15,6,2),('mirror-fresh',15,0,6,2)]:
 for keep in [False,True]:
  name=f'table-keep-{profile}-'+('on' if keep else 'off');x=E.fromstring(parts['word/document.xml']);b=x.find('w:body',N);table=b.find('w:tbl',N);row=table.find('w:tr',N)
  for trpr in row.findall('w:trPr',N):row.remove(trpr)
  for ci,cell in enumerate(row.findall('w:tc',N)):
   for child in list(cell):
    if child.tag!=q('tcPr'):cell.remove(child)
   pr=cell.find('w:tcPr',N)
   for old in pr.findall('w:vAlign',N)+pr.findall('w:vMerge',N):pr.remove(old)
   for border in pr.findall('w:tcBorders/*',N):border.set(q('color'),'FF0000')
   if ci==0:
    if prefix:cell.append(paragraph('P',prefix))
    cell.append(paragraph('K',target,keep));cell.append(paragraph('S',suffix))
   else:
    p=paragraph('R',prefix+target+suffix)
    for t in p.findall('w:r/w:t',N):t.text=t.text.replace('C0','C1')
    cell.append(p)
  for borders in table.findall('w:tblPr/w:tblBorders/*',N):borders.set(q('color'),'FF0000')
  if profile.startswith('mirror-'):
   a,c=row.findall('w:tc',N);left=list(a.findall('w:p',N));right=list(c.findall('w:p',N))
   for p in left:a.remove(p)
   for p in right:c.remove(p)
   for p in right:a.append(p)
   for p in left:c.append(p)
  if body:
   p=paragraph('B',body)
   for t in p.findall('w:r/w:t',N):t.text=t.text.replace('T1','T0')
   b.insert(0,p)
  for r in x.xpath('//w:footnoteReference|//w:lastRenderedPageBreak',namespaces=N):r.getparent().remove(r)
  data=dict(parts);data['word/document.xml']=E.tostring(x,encoding='utf-8',xml_declaration=True,standalone=True)
  with ZipFile(OUT/(name+'-input.docx'),'w',ZIP_DEFLATED) as z:
   for part,content in data.items():z.writestr(part,content)
  cases.append({'name':name,'lineCounts':[[prefix,target,suffix]],'profile':profile,'keep':keep})
(OUT/'probe-cases.json').write_text(json.dumps(cases,indent=2)+'\n',encoding='utf-8')