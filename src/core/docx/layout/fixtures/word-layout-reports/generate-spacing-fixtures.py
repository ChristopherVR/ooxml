"""Writes the generated paragraph spacing fixtures. Word's measurements are in the JSON files
of the same name.

issue-359-contextual-spacing-rows.docx: pairs of Normal paragraphs (space after 200, 1.15 lines)
with contextual spacing on either side, each pair after a single-spaced Sep line.
issue-360-line-spacing-rules.docx: three one-line paragraphs for each line spacing rule, each
group followed by a single-spaced Sep line.
issue-360-triple-spacing-page-fit.docx: 120 one-line paragraphs at triple spacing."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
)
RELS = (
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
)
DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
)
FONT = "Times New Roman"
STYLES = (
    f'<?xml version="1.0" encoding="UTF-8"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
    f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Sep"><w:name w:val="Sep"/>'
    '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>'
)
SECTION = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417"'
    ' w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
)
FLAG = "<w:contextualSpacing/>"


def styles(normal):
    return STYLES.replace(
        '<w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>', normal
    )


def spacing(before=None, after=None):
    attributes = ""
    if before is not None:
        attributes += f' w:before="{before}"'
    if after is not None:
        attributes += f' w:after="{after}"'
    return f"<w:spacing{attributes}/>"


def p(text, properties=""):
    return f'<w:p><w:pPr>{properties}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


CASES = [
    ("row1", "", FLAG),
    ("row2", FLAG, spacing(before=200)),
    ("row3", spacing(after=0), FLAG + spacing(before=200)),
    ("row4", FLAG, FLAG),
    ("r2a0", FLAG + spacing(after=0), spacing(before=200)),
    ("r2a100b300", FLAG + spacing(after=100), spacing(before=300)),
    ("r2a300b100", FLAG + spacing(after=300), spacing(before=100)),
    ("r3a300b200", spacing(after=300), FLAG + spacing(before=200)),
    ("r3a100b300", spacing(after=100), FLAG + spacing(before=300)),
    ("none", "", ""),
    ("noneb300", "", spacing(before=300)),
]



def write(name, body, styles_xml):
    document = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{body}{SECTION}</w:body></w:document>'
    with ZipFile(Path(__file__).with_name(f"{name}.docx"), "w", ZIP_DEFLATED) as package:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", RELS),
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/styles.xml", styles_xml),
            ("word/document.xml", document),
        ):
            entry = ZipInfo(part, (1980, 1, 1, 0, 0, 0))
            entry.compress_type = ZIP_DEFLATED
            package.writestr(entry, data)


SEP = '<w:pStyle w:val="Sep"/>'
write(
    "issue-359-contextual-spacing-rows",
    "".join(p(f"{name}S", SEP) + p(f"{name}U", upper) + p(f"{name}L", lower) for name, upper, lower in CASES)
    + p("endS", SEP),
    STYLES,
)

RULES = [
    ("single", 240, "auto"),
    ("m115", 276, "auto"),
    ("m150", 360, "auto"),
    ("m200", 480, "auto"),
    ("m080", 192, "auto"),
    ("exact20", 400, "exact"),
    ("least20", 400, "atLeast"),
    ("least10", 200, "atLeast"),
]
single = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
write(
    "issue-360-line-spacing-rules",
    "".join(
        "".join(
            p(f"{name}{index}", f'<w:spacing w:after="0" w:line="{line}" w:lineRule="{rule}"/>')
            for index in range(3)
        )
        + p(f"{name}s", SEP)
        for name, line, rule in RULES
    ),
    styles(single),
)
write(
    "issue-360-triple-spacing-page-fit",
    "".join(
        p(f"fit300-{index}", '<w:spacing w:after="0" w:line="720" w:lineRule="auto"/>')
        for index in range(120)
    ),
    styles("<w:pPr></w:pPr>"),
)
