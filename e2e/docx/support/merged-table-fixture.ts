import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const borders =
	'<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>';
const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string, width: number, extra = '') =>
	`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${content}</w:tc>`;
const grid = (...columns: number[]) =>
	`<w:tblGrid>${columns.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid>`;

/**
 * Four columns and three rows with a horizontal merge (A1 spans two columns) and a vertical merge
 * (B1 spans rows two and three).
 */
const merged = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>${borders}</w:tblPr>${grid(2000, 2000, 2000, 2000)}
<w:tr>${cell(para('A1'), 4000, '<w:gridSpan w:val="2"/>')}${cell(para('A3'), 2000)}${cell(para('A4'), 2000)}</w:tr>
<w:tr>${cell(para('B1'), 2000, '<w:vMerge w:val="restart"/>')}${cell(para('B2'), 2000)}${cell(para('B3'), 2000)}${cell(para('B4'), 2000)}</w:tr>
<w:tr>${cell(para('C1'), 2000, '<w:vMerge/>')}${cell(para('C2'), 2000)}${cell(para('C3'), 2000)}${cell(para('C4'), 2000)}</w:tr></w:tbl>`;

/** A one-row table whose first cell holds a two-by-two table. */
const nested = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${borders}</w:tblPr>${grid(2000, 2000)}
<w:tr>${cell(para('N11'), 2000)}${cell(para('N12'), 2000)}</w:tr>
<w:tr>${cell(para('N21'), 2000)}${cell(para('N22'), 2000)}</w:tr></w:tbl>`;
const outer = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>${borders}</w:tblPr>${grid(5000, 3000)}
<w:tr>${cell(para('Outer') + nested + para('Tail'), 5000)}${cell(para('Next'), 3000)}</w:tr></w:tbl>`;

/** A small document with a merged table, a gap paragraph and a table with a nested table. */
export async function mergedTableFixture(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${merged}${para('Gap')}${outer}${para('After')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}
