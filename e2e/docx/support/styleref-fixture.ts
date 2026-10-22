// A two-page document whose header repeats a document number (a paragraph style named "Doc Number"
// whose id is DocNumber) and the page's heading through STYLEREF fields.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';

const paragraph = (text: string, style?: string) =>
	`<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const field = (instr: string, result: string) =>
	`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
	`<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

export async function styleRefFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${r}/header" Target="header1.xml"/></Relationships>`,
	);
	zip.file(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="DocNumber"><w:name w:val="Doc Number"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>${paragraph('DOC-0001', 'DocNumber')}${paragraph('Scope', 'Heading1')}${paragraph('What the plan covers.')}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${paragraph('Method', 'Heading1')}${paragraph('How the work is done.')}<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}"><w:p><w:r><w:t xml:space="preserve">No. </w:t></w:r>${field('STYLEREF "Doc Number" \\* MERGEFORMAT', 'DOC-0001')}<w:r><w:t xml:space="preserve"> | </w:t></w:r>${field('STYLEREF 1', 'Scope')}</w:p></w:hdr>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/**
 * A heading with a line break, and body STYLEREF fields: one whose result is already that heading,
 * one with an old result.
 */
export async function styleRefLineBreakFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${r}/styles" Target="styles.xml"/></Relationships>`,
	);
	zip.file(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`,
	);
	const heading = '<w:t>Chapter 1</w:t><w:br/><w:t>Introduction</w:t>';
	const styleRef = (result: string) =>
		`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> STYLEREF 1 </w:instrText></w:r>` +
		`<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>${result}</w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
	const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r>${heading}</w:r></w:p><w:p>${text('See ')}${styleRef(heading)}${text('.')}</w:p><w:p>${text('Then ')}${styleRef('<w:t>old</w:t>')}${text('.')}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
