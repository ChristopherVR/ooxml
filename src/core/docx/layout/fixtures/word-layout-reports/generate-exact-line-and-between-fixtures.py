"""Writes the fixtures of issues #419 (exact line spacing baseline) and #417 (paragraph border with
only w:between), rebuilt from the XML the reports quote (the reporter's own files are not
published). Word's measurements are in the issues.

exact-line-baseline: one-line paragraphs with exact line heights in Arial and Times New Roman.
single-spacing-control: the same lines with single spacing.
between-mid-page, between-at-page-end, between-not-shared-control: see issue #417."""

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
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1417" w:bottom="1418"'
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




def para(text, font="Arial", size=20, line=None, border=None):
    ppr = ""
    if line:
        ppr += f'<w:spacing w:after="0" w:line="{line}" w:lineRule="exact"/>'
    else:
        ppr += '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
    if border:
        ppr = f"<w:pBdr>{border}</w:pBdr>" + ppr
    return f"<w:p><w:pPr>{ppr}</w:pPr>{run(text, font, size)}</w:p>"


LINES = (
    ("Arial", 20, 200), ("Arial", 20, 300), ("Arial", 20, 400),
    ("Arial", 16, 300), ("Arial", 40, 300), ("Times New Roman", 20, 300),
)
for name, exact in (("exact-line-baseline", True), ("single-spacing-control", False)):
    body = "".join(
        para(f"Line {i} {font} {size // 2} pt", font, size, line if exact else None)
        for i, (font, size, line) in enumerate(LINES, 1)
    )
    write(name, body, styles("Arial", 20), numbering("&#x2022;", "Arial"))

BETWEEN = lambda space: f'<w:between w:val="single" w:sz="4" w:space="{space}" w:color="auto"/>'
write(
    "between-mid-page",
    para("HEADING", border=BETWEEN(15)) + para("Next paragraph", border=BETWEEN(15)),
    styles("Arial", 20), numbering("&#x2022;", "Arial"),
)
FILLER = "".join(para(f"Filler {i}", line=280) for i in range(48))
for name, second in (
    ("between-at-page-end", BETWEEN(15)),
    ("between-not-shared-control", None),
):
    write(
        name,
        FILLER + para("HEADING", line=280, border=BETWEEN(15))
        + para("Next paragraph", line=280, border=second),
        styles("Arial", 20), numbering("&#x2022;", "Arial"),
    )
