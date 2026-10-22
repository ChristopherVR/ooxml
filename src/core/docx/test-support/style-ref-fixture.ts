import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** Paragraph styles whose names differ from their ids, a heading and a character style. */
export const STYLE_REF_STYLES = `<w:styles xmlns:w="${w}">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
<w:style w:type="paragraph" w:styleId="DocNumber"><w:name w:val="Doc Number"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
<w:style w:type="character" w:styleId="EntryName"><w:name w:val="Entry Name"/></w:style>
</w:styles>`;

/** A paragraph, optionally in a style, holding `inner` (runs) or one run of `text`. */
export const paragraph = (text: string, style?: string, inner?: string) =>
	`<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${inner ?? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`}</w:p>`;

/** A complex STYLEREF field whose cached result is `result`, one run per entry. */
export const complexField = (
	instr: string,
	result: string[],
	begin = '<w:fldChar w:fldCharType="begin"/>',
) =>
	`<w:r>${begin}</w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
	result
		.map(
			(text, index) =>
				`<w:r><w:rPr>${index ? '' : '<w:b/>'}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`,
		)
		.join('') +
	'<w:r><w:fldChar w:fldCharType="end"/></w:r>';

export const simpleField = (instr: string, result: string) =>
	`<w:fldSimple w:instr=" ${instr} "><w:r><w:t xml:space="preserve">${result}</w:t></w:r></w:fldSimple>`;

/** A one-section package with a default header and footer around `body`. */
export async function styleRefFixture(parts: {
	body: string;
	header?: string;
	footer?: string;
	styles?: string;
	numbering?: string;
}): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${r}/header" Target="header1.xml"/><Relationship Id="rId3" Type="${r}/footer" Target="footer1.xml"/>${parts.numbering ? `<Relationship Id="rId4" Type="${r}/numbering" Target="numbering.xml"/>` : ''}</Relationships>`,
	);
	zip.file('word/styles.xml', parts.styles ?? STYLE_REF_STYLES);
	if (parts.numbering) zip.file('word/numbering.xml', parts.numbering);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>${parts.body}<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}">${parts.header ?? paragraph('Header')}</w:hdr>`,
	);
	zip.file(
		'word/footer1.xml',
		`<w:ftr xmlns:w="${w}">${parts.footer ?? paragraph('Footer')}</w:ftr>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
