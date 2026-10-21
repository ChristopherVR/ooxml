// A form template's table of sources: each row's kind is a drop-down list content control around
// its cell (`w:tr > w:sdt > w:tc`), the first one temporary, with its placeholder text in the
// glossary part; a check box control around the last row's cell.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PLACEHOLDER = 'Pick a kind.';

const run = (value: string, style = '') =>
	`<w:r>${style ? `<w:rPr><w:rStyle w:val="${style}"/></w:rPr>` : ''}<w:t xml:space="preserve">${value}</w:t></w:r>`;
const cell = (content: string) => `<w:tc><w:p>${content}</w:p></w:tc>`;
const kind = (id: string, temporary: boolean) =>
	`<w:sdt><w:sdtPr><w:alias w:val="Kind ${id}"/><w:id w:val="${id}"/>` +
	`<w:placeholder><w:docPart w:val="KindPlaceholder"/></w:placeholder>${temporary ? '<w:temporary/>' : ''}` +
	`<w:showingPlcHdr/><w:dropDownList><w:listItem w:displayText="Primary" w:value="primary"/>` +
	`<w:listItem w:displayText="Secondary" w:value="secondary"/></w:dropDownList></w:sdtPr>` +
	`<w:sdtContent>${cell(run(PLACEHOLDER, 'PlaceholderText'))}</w:sdtContent></w:sdt>`;
const box =
	`<w:sdt><w:sdtPr><w:alias w:val="Approved"/><w:id w:val="83"/><w14:checkbox><w14:checked w14:val="0"/>` +
	`<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>` +
	`</w14:checkbox></w:sdtPr><w:sdtContent>${cell(run('☐'))}</w:sdtContent></w:sdt>`;

export async function cellControlsFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/glossary/document.xml',
		`<w:glossaryDocument xmlns:w="${w}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="KindPlaceholder"/>` +
			`</w:docPartPr><w:docPartBody><w:p>${run(PLACEHOLDER, 'PlaceholderText')}</w:p></w:docPartBody></w:docPart>` +
			`</w:docParts></w:glossaryDocument>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:w14="${w14}"><w:body><w:p>${run('Sources')}</w:p><w:tbl>` +
			`<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>` +
			`<w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
			`<w:tr>${cell(run('First source'))}${kind('81', true)}</w:tr>` +
			`<w:tr>${cell(run('Second source'))}${kind('82', false)}</w:tr>` +
			`<w:tr>${cell(run('Approved'))}${box}</w:tr>` +
			`</w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
