// Content controls an edit changes as a whole: a drop-down list around a table cell whose paragraph
// sits in a block control that may not be edited, a temporary inline control, and a temporary
// block control around a paragraph, a table and a paragraph.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const run = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const paragraph = (content: string) => `<w:p>${content}</w:p>`;
const sdt = (properties: string, content: string) =>
	`<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const table = (widths: number[], cells: string) =>
	`<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid>` +
	`${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`;
const cell = (content: string) => `<w:tc>${content}</w:tc>`;

const kind = sdt(
	`<w:alias w:val="Kind"/><w:id w:val="91"/><w:dropDownList>` +
		`<w:listItem w:displayText="Primary" w:value="primary"/>` +
		`<w:listItem w:displayText="Secondary" w:value="secondary"/></w:dropDownList>`,
	cell(
		sdt(
			`<w:alias w:val="Fixed note"/><w:id w:val="92"/><w:lock w:val="contentLocked"/>`,
			paragraph(run('Fixed')),
		),
	),
);
const draft = sdt(
	`<w:alias w:val="Draft"/><w:id w:val="93"/><w:temporary/><w:text/>`,
	run('Draft'),
);
const section = sdt(
	`<w:alias w:val="Section"/><w:id w:val="94"/><w:temporary/>`,
	paragraph(run('Before')) +
		table([9000], cell(paragraph(run('Inside')))) +
		paragraph(run('After')),
);

export async function controlEditsFixture(): Promise<Uint8Array> {
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
		`<w:document xmlns:w="${w}"><w:body>${paragraph(run('Edits'))}` +
			`${table([4500, 4500], cell(paragraph(run('Kind'))) + kind)}` +
			`${paragraph(run('Note: ') + draft)}${section}${paragraph(run('End'))}` +
			`<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
