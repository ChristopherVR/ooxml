from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def p(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def sdt(id, title, content):
    return (
        f'<w:sdt><w:sdtPr><w:alias w:val="{title}"/><w15:appearance w15:val="tags"/>'
        f'<w:id w:val="{id}"/></w:sdtPr>'
        f"<w:sdtContent>{content}</w:sdtContent></w:sdt>"
    )


def cell(text):
    return f"<w:tc>{p(text)}</w:tc>"


block = sdt(1, "Block", p("Lorem ipsum") + p("dolor sit") + p("amet"))
cells = sdt(3, "Cells", cell("Consectetur") + cell("adipiscing"))
rows = sdt(2, "Rows", f"<w:tr>{cells}</w:tr><w:tr>{cell('elit')}{cell('eiusmod')}</w:tr>")
table = (
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>'
    f"</w:tblGrid>{rows}</w:tbl>"
)
parts = {
    "[Content_Types].xml": (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    ),
    "word/document.xml": (
        f'<w:document xmlns:w="{W}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15">'
        f"<w:body>{block}{table}{p('Tempor')}<w:sectPr/></w:body></w:document>"
    ),
}

with ZipFile(Path(__file__).with_name("block-content-control-tags.docx"), "w") as docx:
    for path, xml in parts.items():
        docx.writestr(ZipInfo(path), xml, compress_type=ZIP_DEFLATED)
