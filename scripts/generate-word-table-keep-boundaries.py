import sys,json
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
root=Path(__file__).resolve().parent.parent;out=Path(sys.argv[1]);out.mkdir(parents=True,exist_ok=True)
w='http://schemas.openxmlformats.org/wordprocessingml/2006/main';ns={'w':w};q=lambda n:'{'+w+'}'+n
records=[]
for border,padding,mirror in [(3,0,False),(3,0,True),(6,2,False),(1.5,1.5,False)]:
 for delta in [-1,0,1]:
  for keep in [False,True]:
   template='table-keep-'+('mirror-fresh' if mirror else 'fresh')+'-'+('on' if keep else 'off')+'.docx'
   with ZipFile(root/'src/core/docx/__fixtures__/table-keep-lines'/template) as z:parts={n:z.read(n) for n in z.namelist()}
   doc=E.fromstring(parts['word/document.xml']);body=doc.find('w:body',ns);p=body.find('w:p',ns)
   runs=p.findall('w:r',ns)
   for run in runs[1:]:p.remove(run)
   minimum=72+2*border+2*padding;remaining=minimum+delta;occupied=220-remaining
   p.find('w:pPr/w:spacing',ns).set(q('line'),str(round(occupied*20)))
   table=body.find('w:tbl',ns)
   for edge in table.findall('.//w:tcBorders/*',ns)+table.findall('w:tblPr/w:tblBorders/*',ns):edge.set(q('sz'),str(round(border*8)))
   for cell in table.findall('w:tr/w:tc',ns):
    pr=cell.find('w:tcPr',ns);margin=pr.find('w:tcMar',ns)
    if margin is None:margin=E.SubElement(pr,q('tcMar'))
    for side in ['top','bottom']:
     item=margin.find('w:'+side,ns)
     if item is None:item=E.SubElement(margin,q(side))
     item.set(q('w'),str(round(padding*20)));item.set(q('type'),'dxa')
   name=f'keep-boundary-b{border:g}-p{padding:g}-'+('right' if mirror else 'left')+f'-d{delta}-'+('on' if keep else 'off')
   parts['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True,standalone=True)
   with ZipFile(out/(name+'-input.docx'),'w',ZIP_DEFLATED) as z:
    for namepart,data in parts.items():z.writestr(namepart,data)
   records.append({'name':name,'borderPt':border,'paddingPt':padding,'remainingPt':remaining,'minimumPt':minimum,'keep':keep,'mirror':mirror})
(out/'probe-cases.json').write_text(json.dumps(records,indent=2)+'\n',encoding='utf8')