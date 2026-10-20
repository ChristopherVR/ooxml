// A form template with legacy form fields: a text field in a body paragraph, a check box and a
// drop-down in a label/value table. `protect` adds Word's "Filling in forms" protection.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const run = (inner: string) => `<w:r>${inner}</w:r>`;
const text = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);

function formField(name: string, data: string, code: string, result: string | null, id: number) {
	return (
		`<w:bookmarkStart w:id="${id}" w:name="${name}"/>` +
		run(
			`<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/><w:enabled/><w:calcOnExit w:val="0"/>${data}</w:ffData></w:fldChar>`,
		) +
		run(`<w:instrText xml:space="preserve"> ${code} </w:instrText>`) +
		(result === null ? '' : run('<w:fldChar w:fldCharType="separate"/>') + text(result)) +
		run('<w:fldChar w:fldCharType="end"/>') +
		`<w:bookmarkEnd w:id="${id}"/>`
	);
}

const cell = (content: string) => `<w:tc><w:p>${content}</w:p></w:tc>`;

export async function formFieldsFixture(protect = false): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/settings" Target="settings.xml"/></Relationships>`,
	);
	zip.file(
		'word/settings.xml',
		`<w:settings xmlns:w="${w}">${protect ? '<w:documentProtection w:edit="forms" w:enforcement="1"/>' : ''}</w:settings>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>` +
			`<w:p>${text('Name: ')}${formField('Name', '<w:textInput><w:default w:val="Name"/><w:maxLength w:val="20"/></w:textInput>', 'FORMTEXT', 'Name', 1)}</w:p>` +
			`<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>` +
			`<w:tr>${cell(text('Approved'))}${cell(formField('Approved', '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>', 'FORMCHECKBOX', null, 2))}</w:tr>` +
			`<w:tr>${cell(text('Status'))}${cell(formField('Status', '<w:ddList><w:result w:val="1"/><w:listEntry w:val="New"/><w:listEntry w:val="Revised"/><w:listEntry w:val="Withdrawn"/></w:ddList>', 'FORMDROPDOWN', null, 3))}</w:tr>` +
			`</w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
