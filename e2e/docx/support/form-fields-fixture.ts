// A form template with legacy form fields: a text field in a body paragraph, a check box and a
// drop-down in a label/value table. `protect` adds Word's "Filling in forms" protection.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const run = (inner: string) => `<w:r>${inner}</w:r>`;
const text = (value: string, props = '') =>
	run(`${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${value}</w:t>`);

/** A legacy form field; its result is plain text, or the result's runs as written. */
function formField(
	name: string,
	data: string,
	code: string,
	result: string | { runs: string } | null,
	id: number,
) {
	return (
		`<w:bookmarkStart w:id="${id}" w:name="${name}"/>` +
		run(
			`<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/><w:enabled/><w:calcOnExit w:val="0"/>${data}</w:ffData></w:fldChar>`,
		) +
		run(`<w:instrText xml:space="preserve"> ${code} </w:instrText>`) +
		(result === null
			? ''
			: run('<w:fldChar w:fldCharType="separate"/>') +
				(typeof result === 'string' ? text(result) : result.runs)) +
		run('<w:fldChar w:fldCharType="end"/>') +
		`<w:bookmarkEnd w:id="${id}"/>`
	);
}

const cell = (content: string) => `<w:tc><w:p>${content}</w:p></w:tc>`;

const checkBox = '<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>';
const FORMS_PROTECTION = '<w:documentProtection w:edit="forms" w:enforcement="1"/>';

async function formPackage(settings: string, body: string): Promise<Uint8Array> {
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
	zip.file('word/settings.xml', `<w:settings xmlns:w="${w}">${settings}</w:settings>`);
	zip.file('word/document.xml', `<w:document xmlns:w="${w}"><w:body>${body}</w:body></w:document>`);
	return zip.generateAsync({ type: 'uint8array' });
}

export function formFieldsFixture(protect = false): Promise<Uint8Array> {
	return formPackage(
		protect ? FORMS_PROTECTION : '',
		`<w:p>${text('Name: ')}${formField('Name', '<w:textInput><w:default w:val="Name"/><w:maxLength w:val="20"/></w:textInput>', 'FORMTEXT', 'Name', 1)}</w:p>` +
			`<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>` +
			`<w:tr>${cell(text('Approved'))}${cell(formField('Approved', checkBox, 'FORMCHECKBOX', null, 2))}</w:tr>` +
			`<w:tr>${cell(text('Status'))}${cell(formField('Status', '<w:ddList><w:result w:val="1"/><w:listEntry w:val="New"/><w:listEntry w:val="Revised"/><w:listEntry w:val="Withdrawn"/></w:ddList>', 'FORMDROPDOWN', null, 3))}</w:tr>` +
			`</w:tbl><w:sectPr/>`,
	);
}

/**
 * A form protected for filling in with a blank text field of four characters at most, as Word
 * writes a text field that holds no text: five en spaces, one run each, kept from proofing.
 */
export function blankFieldFixture(): Promise<Uint8Array> {
	const blank = { runs: Array.from({ length: 5 }, () => text('\u2002', '<w:noProof/>')).join('') };
	return formPackage(
		FORMS_PROTECTION,
		`<w:p>${text('Code: ')}${formField('Code', '<w:textInput><w:maxLength w:val="4"/></w:textInput>', 'FORMTEXT', blank, 1)}${text(' end')}</w:p>` +
			`<w:sectPr/>`,
	);
}

/** A form protected for filling in whose drop-down lists one entry name twice. */
export function repeatedEntriesFixture(): Promise<Uint8Array> {
	const entries = ['Pass', 'Fail', 'Pass'].map((entry) => `<w:listEntry w:val="${entry}"/>`);
	return formPackage(
		FORMS_PROTECTION,
		`<w:p>${text('Grade: ')}${formField('Grade', `<w:ddList><w:result w:val="0"/>${entries.join('')}</w:ddList>`, 'FORMDROPDOWN', null, 1)}</w:p>` +
			`<w:sectPr/>`,
	);
}

/**
 * A form protected for filling in: a check box, one inside a content control that locks its
 * content, and a last section the protection leaves editable (`w:formProt w:val="0"`).
 */
export function protectedSectionsFixture(): Promise<Uint8Array> {
	return formPackage(
		FORMS_PROTECTION,
		`<w:p>${text('Free ')}${formField('Free', checkBox, 'FORMCHECKBOX', null, 1)}</w:p>` +
			`<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:id w:val="9"/></w:sdtPr><w:sdtContent>` +
			`<w:p>${text('Locked ')}${formField('Locked', checkBox, 'FORMCHECKBOX', null, 2)}</w:p></w:sdtContent></w:sdt>` +
			`<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr>${text('End of form')}</w:p>` +
			`<w:p>${text('Notes')}</w:p>` +
			`<w:sectPr><w:type w:val="continuous"/><w:formProt w:val="0"/></w:sectPr>`,
	);
}
