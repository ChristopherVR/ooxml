"""Writes the generated fixtures of issues #436 and #438 (the reporter's own files were not
available). Reuses the package parts of generate-g5-fixtures.py.

issue-436-*: automatic-colour text on cell and paragraph shading, and the same with an explicit
black colour (control). Word 16 draws the automatic text white on 000000, 371127 and C00000 and
black on 404040, 595959 and 4472C4.
issue-438-*: a last paragraph that ends in a page break (alone, or after text), and the control
with an empty paragraph after the break."""

import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('g5', Path(__file__).with_name('generate-g5-fixtures.py'))
g5 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(g5)

FILLS = ['000000', '371127', 'C00000', '404040', '595959', '4472C4']


def cell(fill, color):
    rpr = f'<w:rPr><w:color w:val="{color}"/></w:rPr>' if color else ''
    return (f'<w:tc><w:tcPr><w:tcW w:w="4536" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r>{rpr}<w:t>Cell {fill}</w:t></w:r></w:p></w:tc>')


def shaded(fill, color):
    rpr = f'<w:rPr><w:color w:val="{color}"/></w:rPr>' if color else ''
    return (f'<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:pPr>'
            f'<w:r>{rpr}<w:t>Paragraph {fill}</w:t></w:r></w:p>')


def build(color):
    rows = ''.join(f'<w:tr>{cell(f, color)}</w:tr>' for f in FILLS)
    table = f'<w:tbl><w:tblPr><w:tblW w:w="4536" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4536"/></w:tblGrid>{rows}</w:tbl>'
    return table + ''.join(shaded(f, color) for f in FILLS[:3] + FILLS[3:4])


g5.write('issue-436-auto-colour-on-dark-shading.docx', build(None))
g5.write('issue-436-auto-colour-on-dark-shading-control.docx', build('000000'))

lorem = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.'
br = '<w:br w:type="page"/>'
g5.write('issue-438-page-break-at-document-end.docx',
         f'<w:p><w:r><w:t>{lorem}</w:t></w:r></w:p><w:p><w:r>{br}</w:r></w:p>')
g5.write('issue-438-page-break-after-text-at-document-end.docx',
         f'<w:p><w:r><w:t>{lorem}</w:t></w:r><w:r>{br}</w:r></w:p>')
g5.write('issue-438-page-break-at-document-end-control.docx',
         f'<w:p><w:r><w:t>{lorem}</w:t></w:r></w:p><w:p><w:r>{br}</w:r></w:p><w:p/>')
