"""Writes the line height fixtures of issues #391, #393 and #395, rebuilt from the XML the reports
quote (the reporter's own files are not published). Word's measurements are in the issues.

mixed-font-line-height(-control): 60 one-line paragraphs in 12 pt Times New Roman with a ", " run in
12 pt Garamond in the middle (control: the comma in Times New Roman).
symbol-bullet-line-height(-control): 28 two-line bullet items in 11 pt Calibri, the bullet in
Symbol (control: a bullet in Calibri).
larger-space-run(-mid-line, -control): 30 two-line paragraphs in 11 pt Calibri with the space a
separate 12 pt bold run (control: 11 pt)."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
)
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL}styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{REL}numbering" Target="numbering.xml"/>'
    f'<Relationship Id="rId3" Type="{REL}settings" Target="settings.xml"/></Relationships>'
)
SETTINGS = (
    f'<?xml version="1.0" encoding="UTF-8"?><w:settings {W}><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
)
SECTION = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417"'
    ' w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
)


def styles(font, size):
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
    )


def numbering(bullet, font):
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><w:numbering {W}><w:abstractNum w:abstractNumId="0">'
        '<w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{bullet}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    )


def run(text, font, size, bold=False):
    b = "<w:b/>" if bold else ""
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>{b}<w:sz w:val="{size}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def write(name, body, styles_xml, numbering_xml):
    document = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{body}{SECTION}</w:body></w:document>'
    with ZipFile(Path(__file__).with_name(f"{name}.docx"), "w", ZIP_DEFLATED) as package:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", RELS),
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/styles.xml", styles_xml),
            ("word/numbering.xml", numbering_xml),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", document),
        ):
            entry = ZipInfo(part, (1980, 1, 1, 0, 0, 0))
            entry.compress_type = ZIP_DEFLATED
            package.writestr(entry, data)


TNR = "Times New Roman"
for name, comma in (("mixed-font-line-height", "Garamond"), ("mixed-font-line-height-control", TNR)):
    body = "".join(
        f"<w:p>{run(f'Line {i} in Times New Roman', TNR, 24)}{run(', ', comma, 24)}{run('then Times again', TNR, 24)}</w:p>"
        for i in range(1, 61)
    )
    write(name, body, styles(TNR, 24), numbering("&#x2022;", TNR))

LONG = " ".join(w * 4 for w in "abcdefghijklmnopqrst")
for name, bullet, font in (
    ("symbol-bullet-line-height", "&#xF0B7;", "Symbol"),
    ("symbol-bullet-line-height-control", "&#x2022;", "Calibri"),
):
    body = "".join(
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f"{run(f'{i} {LONG} {LONG}', 'Calibri', 22)}</w:p>"
        for i in range(1, 29)
    )
    write(name, body, styles("Calibri", 22), numbering(bullet, font))

HEAD = " ".join(w * 4 for w in "abcdefghijklmnopqrst")
for name, size, mid in (
    ("larger-space-run", 24, False),
    ("larger-space-run-mid-line", 24, True),
    ("larger-space-run-control", 22, False),
):
    paragraphs = []
    for i in range(1, 31):
        if mid:
            first = f"{i} aaaa bbbb"
            rest = f"cccc dddd {HEAD} {HEAD} second line text"
        else:
            first = f"{i} {HEAD}"
            rest = f"uuuu vvvv wwww {HEAD} second line text"
        paragraphs.append(
            f"<w:p>{run(first, 'Calibri', 22)}{run(' ', 'Calibri', size, True)}{run(rest, 'Calibri', 22)}</w:p>"
        )
    write(name, "".join(paragraphs), styles("Calibri", 22), numbering("&#x2022;", "Calibri"))
