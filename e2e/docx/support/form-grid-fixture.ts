// A form template's label/value grid: rows of 2, 4, 1 and 2 physical cells over a four-column
// grid, shaded label cells and an inline check box content control.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const text = (value: string) => `<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;
const cell = (content: string, span = 1) =>
	`<w:tc>${span > 1 ? `<w:tcPr><w:gridSpan w:val="${span}"/></w:tcPr>` : ''}${content}</w:tc>`;
const label = (value: string) =>
	`<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr>${text(value)}</w:tc>`;
const checkbox =
	`<w:sdt><w:sdtPr><w:id w:val="7"/><w14:checkbox><w14:checked w14:val="0"/>` +
	`<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>` +
	`</w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>` +
	`<w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;

export const formGridRowCellCounts = [2, 4, 1, 2];

export async function formGridFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:w14="${w14}"><w:body>${text('Transmittal form')}<w:tbl>` +
			`<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>` +
			`<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2500"/><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>` +
			`<w:tr>${label('Document number')}${cell(text('Number'), 3)}</w:tr>` +
			`<w:tr>${label('Revision')}${cell(text('Rev'))}${label('Date')}${cell(text('Date'))}</w:tr>` +
			`<w:tr>${cell(text('Change history'), 4)}</w:tr>` +
			`<w:tr>${label('Approved')}${cell(`<w:p>${checkbox}<w:r><w:t xml:space="preserve"> Yes</w:t></w:r></w:p>`, 3)}</w:tr>` +
			`</w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
