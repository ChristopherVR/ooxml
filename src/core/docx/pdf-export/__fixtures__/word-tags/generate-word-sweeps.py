"""Generate sweep .docx files for issue 415. Usage: gen.py <name>  (sweeps defined below)."""
import sys, zipfile, json
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = ('xmlns:w="%s" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' % W)
_id = [3000]
def nid():
    _id[0] += 1; return _id[0]
def run(t): return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % t
def sdt(alias, body, tags, tag=True):
    ap = '<w15:appearance w15:val="%s"/>' % ('tags' if tags is True else tags) if tags else ''
    return ('<w:sdt><w:sdtPr><w:alias w:val="%s"/><w:tag w:val="%s"/><w:id w:val="%d"/>%s</w:sdtPr>'
            '<w:sdtContent>%s</w:sdtContent></w:sdt>') % (alias, alias, nid(), ap, body)
def para(n, prec='untagged', tagmode='both', jc='both', ind=0, tail=' alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau.', mid=None, lead='i'):
    """prec: untagged | plain | tagged ; tagmode: both | inner | outer | none"""
    p = lead * n + ' lorem ipsum dolor sit amet '
    body = run(p)
    if prec == 'plain': body += run('Name/Field')
    else: body += sdt('Field', run('Name/Field'), prec == 'tagged')
    if tagmode == 'both':
        body += sdt('Outer', sdt('Inner', run('Value'), True) + run(' text'), True)
    elif tagmode == 'inner':
        body += sdt('Inner', run('Value'), True) + run(' text')
    elif tagmode == 'outer':
        body += sdt('Outer', run('Value') + run(' text'), True)
    else:
        body += run('Value text')
    body += run(tail)
    pr = '<w:jc w:val="%s"/>' % jc
    if ind: pr = '<w:ind w:right="%d"/>' % ind + pr
    return '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>%s</w:pPr>%s</w:p>' % (pr, body)
def pagebreak_para(inner): return inner
def write(name, paras, font='Arial'):
    body = ''.join(paras)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>') % (NS, body)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="%s"><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % ((W, font, font, font, font))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="%s"><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    with zipfile.ZipFile(name + '.docx', 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', settings); z.writestr('word/_rels/document.xml.rels', drels)

SW = {}
# sweep A: reporter's sweep, i count 15..40, untagged vs plain, both jc
SW['A'] = [dict(n=n, prec=p, jc=j) for p in ('untagged', 'plain') for j in ('both', 'left') for n in range(15, 41)]
combos = [('untagged','both','left'),('untagged','inner','left'),('untagged','none','left'),('untagged','outer','left'),('tagged','both','left'),('plain','both','left'),('untagged','both','both'),('plain','both','both')]
SW['B'] = [dict(n=28, prec=p, tagmode=t, jc=j, ind=r) for (p,t,j) in combos for r in range(0,121)]
if __name__ == '__main__':
    name = sys.argv[1]
    specs = SW[name]
    write(name, [para(**s) for s in specs])
    json.dump(specs, open(name + '.json', 'w'))
    print(len(specs), 'paragraphs')
