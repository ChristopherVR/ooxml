"""Writes the generated fixtures of issues #411, #390, #389, #388 and #403 (A4, 1417 twip
margins, 14 pt exact filler lines, so the page body is 700.2 pt high). Word's measurements are in
the tests next to the layout code.

issue-411-*: page breaks that share a run with text, and the same breaks in runs of their own.
issue-390-*: a keepNext heading after 48 filler lines, with and without keepLines.
issue-389-*: a shaded box-bordered heading after 47 filler lines, with and without a bottom border.
issue-388-*: one top-bordered paragraph per compound style and size, each after a plain line.
issue-403-*: top borders in theme colours with a shade and a tint, and the same in plain colours."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = (
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
)
DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'
)
STYLES = (
    f'<?xml version="1.0" encoding="UTF-8"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>'
)
FILL = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
LINE = f'<a:ln w="9525">{FILL}</a:ln>'
EFFECT = '<a:effectStyle><a:effectLst/></a:effectStyle>'
THEME = (
    '<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
    '<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
    '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
    '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
    '<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
    '<a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    f'<a:fmtScheme name="Office"><a:fillStyleLst>{FILL * 3}</a:fillStyleLst>'
    f'<a:lnStyleLst>{LINE * 3}</a:lnStyleLst><a:effectStyleLst>{EFFECT * 3}</a:effectStyleLst>'
    f'<a:bgFillStyleLst>{FILL * 3}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'
)
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417"'
    ' w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
)
EXACT = '<w:spacing w:after="0" w:line="280" w:lineRule="exact"/>'


def p(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def fill(count):
    return ''.join(p(f'Filler {i + 1}', EXACT) for i in range(count))


def write(name, body):
    path = Path(__file__).parent / name
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        parts = {
            '[Content_Types].xml': CT,
            '_rels/.rels': RELS,
            'word/_rels/document.xml.rels': DOC_RELS,
            'word/document.xml': f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>',
            'word/styles.xml': STYLES,
            'word/theme/theme1.xml': THEME,
        }
        for part, data in parts.items():
            info = ZipInfo(part, (2026, 11, 28, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)


br = '<w:br w:type="page"/>'
write('issue-411-page-break-in-text-run.docx',
      '<w:p><w:r><w:t>Page one.</w:t></w:r></w:p>'
      f'<w:p><w:r>{br}<w:t>Page two: the break and this text are in the same run.</w:t></w:r></w:p>'
      f'<w:p><w:r><w:t>Still page two, then a break after text in the same run.</w:t>{br}</w:r></w:p>'
      '<w:p><w:r><w:t>Page three.</w:t></w:r></w:p>')
write('issue-411-page-break-own-run.docx',
      '<w:p><w:r><w:t>Page one.</w:t></w:r></w:p>'
      f'<w:p><w:r>{br}</w:r><w:r><w:t>Page two: the break and this text are in separate runs.</w:t></w:r></w:p>'
      f'<w:p><w:r><w:t>Still page two, then a break after text in its own run.</w:t></w:r><w:r>{br}</w:r></w:p>'
      '<w:p><w:r><w:t>Page three.</w:t></w:r></w:p>')


def heading(extra):
    return p('Heading', f'<w:keepNext/>{extra}<w:spacing w:after="240" w:line="280" w:lineRule="exact"/>')


tail = ''.join(p(f'Body {i + 1}', EXACT) for i in range(3))
write('issue-390-keepnext-spacing-after.docx', fill(48) + heading('') + tail)
write('issue-390-keepnext-spacing-after-keeplines-control.docx', fill(48) + heading('<w:keepLines/>') + tail)


def box(bottom):
    def side(name, space):
        return f'<w:{name} w:val="single" w:sz="4" w:space="{space}" w:color="auto"/>'

    return ('<w:pBdr>' + side('top', 15) + side('left', 4) + (side('bottom', 15) if bottom else '')
            + side('right', 4) + '</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' + EXACT)


write('issue-389-border-bottom-space-page-end.docx', fill(47) + p('Heading', box(True)) + tail)
write('issue-389-border-bottom-space-page-end-control.docx', fill(47) + p('Heading', box(False)) + tail)

STYLES_388 = ['single', 'double', 'thinThickSmallGap', 'thickThinSmallGap', 'thinThickMediumGap',
              'thinThickLargeGap', 'thinThickThinSmallGap']


def compound(style, size):
    return p(f'{style} {size}', f'<w:pBdr><w:top w:val="{style}" w:sz="{size}" w:space="1" w:color="auto"/></w:pBdr>')


for name, styles in (('compound-top-borders', STYLES_388), ('single-top-borders', ['single'] * len(STYLES_388))):
    body = ''
    for size in (12, 24):
        for style in styles:
            body += p('Plain line') + compound(style, size)
    write(f'issue-388-{name}.docx', body)


def themed(first, second):
    def top(attrs):
        return f'<w:pBdr><w:top w:val="single" w:sz="24" w:space="1" {attrs}/></w:pBdr>'

    return p('Plain line') + p('Shade', top(first)) + p('Tint', top(second))


write('issue-403-border-theme-shade-tint.docx',
      themed('w:color="622423" w:themeColor="accent2" w:themeShade="7F"',
             'w:color="E5B8B7" w:themeColor="accent2" w:themeTint="66"'))
write('issue-403-border-plain-colors.docx', themed('w:color="622423"', 'w:color="E5B8B7"'))
