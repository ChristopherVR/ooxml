import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const cell = (text: string) =>
	`<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

/** Three pages: a header and footer, a table with a tall row on page 1, plain text after page breaks. */
export async function multiPageDocx(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${rel}/footer" Target="footer1.xml"/></Relationships>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
	);
	zip.file(
		'word/footer1.xml',
		`<w:ftr xmlns:w="${w}"><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>`,
	);
	const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
	const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>${para('Page one intro')}<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2400" w:hRule="atLeast"/></w:trPr>${cell('Cell A')}${cell('Cell B')}</w:tr></w:tbl>${para('After table')}${pageBreak}${para('Page two first')}${para('Page two second')}${pageBreak}${para('Page three first')}<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}
