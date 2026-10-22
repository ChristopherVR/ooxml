// A synthetic request form, built the way Word writes form templates: a label/value grid with
// shaded label cells and merged value cells, legacy form fields (text, date, check box, drop-down)
// some sharing a name, check box content controls with their labels after them, a full-width
// notes cell, a body paragraph holding only a field, a references table whose last column is a
// cell-level drop-down per row, inline text, date and combo box controls, and a header that
// repeats the document number with STYLEREF.
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

export const ENSPACES = '\u2002'.repeat(5);
export const BALLOT = '\u2610';

const run = (inner: string, props = '') =>
	`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}${inner}</w:r>`;
export const textRun = (value: string, props = '') =>
	run(`<w:t xml:space="preserve">${value}</w:t>`, props);
export const paragraph = (inner: string, style?: string) =>
	`<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${inner}</w:p>`;

let bookmark = 0;
/** A legacy form field as Word writes it: begin with its data, code, result runs, end. */
export function legacyField(
	code: 'FORMTEXT' | 'FORMCHECKBOX' | 'FORMDROPDOWN',
	data: string,
	options: { name?: string; result?: string[]; enabled?: boolean } = {},
): string {
	const name = options.name ?? '';
	const id = ++bookmark;
	const begin = run(
		`<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/>` +
			`${options.enabled === false ? '<w:enabled w:val="0"/>' : '<w:enabled/>'}<w:calcOnExit w:val="0"/>${data}</w:ffData></w:fldChar>`,
	);
	const result = options.result
		? run('<w:fldChar w:fldCharType="separate"/>') +
			options.result.map((text) => textRun(text, '<w:noProof/>')).join('')
		: '';
	const field =
		begin +
		run(`<w:instrText xml:space="preserve"> ${code} </w:instrText>`) +
		result +
		run('<w:fldChar w:fldCharType="end"/>');
	return name
		? `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${field}<w:bookmarkEnd w:id="${id}"/>`
		: field;
}

/**
 * A FORMTEXT field whose default text and result are `text` (Word writes five en spaces as the
 * result of an empty one), with its text input settings in schema order.
 */
export function legacyText(
	text: string,
	options: {
		name?: string;
		enabled?: boolean;
		type?: string;
		maxLength?: number;
		format?: string;
	} = {},
): string {
	const { type, maxLength, format, ...field } = options;
	const input =
		(type ? `<w:type w:val="${type}"/>` : '') +
		(text ? `<w:default w:val="${text}"/>` : '') +
		(maxLength ? `<w:maxLength w:val="${maxLength}"/>` : '') +
		(format ? `<w:format w:val="${format}"/>` : '');
	return legacyField('FORMTEXT', input ? `<w:textInput>${input}</w:textInput>` : '<w:textInput/>', {
		...field,
		result: [text || ENSPACES],
	});
}

/** A content control around runs. */
export const inlineControl = (properties: string, content: string) =>
	`<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** A check box content control showing the empty ballot box. */
export const checkboxControl = (id: number) =>
	inlineControl(
		`<w:id w:val="${id}"/><w14:checkbox><w14:checked w14:val="0"/>` +
			'<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>',
		textRun(BALLOT, '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>'),
	);

const placeholderRun = (text: string) => textRun(text, '<w:rStyle w:val="PlaceholderText"/>');

/** A cell, optionally shaded as a label cell and spanning grid columns. */
export function cell(content: string, options: { span?: number; label?: boolean } = {}): string {
	const properties =
		(options.span ? `<w:gridSpan w:val="${options.span}"/>` : '') +
		(options.label ? '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' : '');
	return `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ''}${content || '<w:p/>'}</w:tc>`;
}
export const label = (text: string, span?: number) =>
	cell(paragraph(textRun(text)), { label: true, ...(span ? { span } : {}) });
const fieldCell = (field: string, span?: number, style?: string) =>
	cell(paragraph(field, style), span ? { span } : {});
export const table = (grid: number[], rows: string[]) =>
	`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid
		.map((width) => `<w:gridCol w:w="${width}"/>`)
		.join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
export const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const headerRow = (...cells: string[]) =>
	`<w:tr><w:trPr><w:tblHeader/></w:trPr>${cells.join('')}</w:tr>`;

/** A drop-down content control around a whole cell, showing its placeholder. */
const dropDownCell = (id: number) =>
	`<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:temporary/><w:showingPlcHdr/><w:dropDownList>` +
	'<w:listItem w:displayText="Reference" w:value="Reference"/><w:listItem w:displayText="Cross reference" w:value="Cross reference"/>' +
	`</w:dropDownList></w:sdtPr><w:sdtContent>${cell(paragraph(placeholderRun('Choose an item.')))}</w:sdtContent></w:sdt>`;

const STATUS_ENTRIES =
	'<w:ddList><w:listEntry w:val="New"/><w:listEntry w:val="Revised"/><w:listEntry w:val="Withdrawn"/></w:ddList>';

/** The form's body. */
export function formTemplateBody(): string {
	bookmark = 0;
	const grid = table(
		[1400, 1400, 1400, 1400, 1400, 1400],
		[
			row(
				label('Project'),
				fieldCell(legacyText('Project'), 3),
				label('Date'),
				fieldCell(legacyText('', { name: 'IssueDate', type: 'date', format: 'dd-MMM-yyyy' })),
			),
			row(
				label('Document No.'),
				fieldCell(legacyText('Number', { maxLength: 12 }), 3, 'DocNumber'),
				label('Rev'),
				fieldCell(legacyText('Rev', { format: 'UPPERCASE' })),
			),
			row(
				label('To'),
				fieldCell(legacyText('Name')),
				fieldCell(legacyText('Name')),
				fieldCell(legacyText('Name')),
				fieldCell(legacyText('Name'), 2),
			),
			row(
				label('Preparer'),
				fieldCell(legacyText('Name')),
				label('Signature'),
				cell(''),
				label('Date'),
				fieldCell(legacyText('', { name: 'Text5' })),
			),
			row(
				label('Reviewer (optional)'),
				fieldCell(legacyText('Name')),
				label('Signature'),
				cell(''),
				label('Date'),
				fieldCell(legacyText('', { name: 'Text5' })),
			),
			row(
				label('Status'),
				fieldCell(legacyField('FORMDROPDOWN', STATUS_ENTRIES, { name: 'Status' }), 5),
			),
			row(
				label('Approved'),
				fieldCell(
					legacyField(
						'FORMCHECKBOX',
						'<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>',
						{ name: 'Approved' },
					),
					2,
				),
				label('Count'),
				fieldCell(legacyText('', { name: 'Count', type: 'number', format: '#,##0.00' }), 2),
			),
			row(label('Archived'), fieldCell(legacyText('No', { name: 'Archived', enabled: false }), 5)),
		],
	);
	const status = table(
		[4200, 4200],
		[
			row(cell(paragraph(textRun('Status of information')), { span: 2, label: true })),
			row(
				label('Needs further checking'),
				cell(
					paragraph(
						`${checkboxControl(101)}${textRun(' Yes ')}${checkboxControl(102)}${textRun(' No')}`,
					),
				),
			),
			row(fieldCell(legacyText(''), 2)),
		],
	);
	const references = table(
		[4200, 2100, 2100],
		[
			headerRow(label('Document title'), label('Revision'), label('Kind')),
			row(cell(''), cell(''), dropDownCell(201)),
			row(cell(''), cell(''), dropDownCell(202)),
		],
	);
	const controls =
		paragraph(
			textRun('Client: ') +
				inlineControl(
					'<w:alias w:val="Client"/><w:tag w:val="client"/><w:id w:val="301"/><w:showingPlcHdr/><w:text/>',
					placeholderRun('Click or tap here to enter text.'),
				),
		) +
		paragraph(
			textRun('Due: ') +
				inlineControl(
					'<w:id w:val="302"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="dd MMMM yyyy"/><w:lid w:val="en-US"/></w:date>',
					placeholderRun('Click or tap to enter a date.'),
				),
		) +
		paragraph(
			textRun('Region: ') +
				inlineControl(
					'<w:tag w:val="region"/><w:id w:val="303"/><w:comboBox><w:listItem w:displayText="North" w:value="N"/><w:listItem w:displayText="South" w:value="S"/></w:comboBox>',
					textRun('North'),
				),
		) +
		paragraph(
			textRun('Owner: ') +
				inlineControl(
					'<w:id w:val="304"/><w:lock w:val="contentLocked"/><w:text/>',
					textRun('Records office'),
				),
		);
	return (
		paragraph(textRun('Request form'), 'Title') +
		grid +
		status +
		paragraph(textRun('Purpose of issue')) +
		paragraph(legacyText('')) +
		references +
		controls
	);
}

const STYLES = `<w:styles xmlns:w="${W}">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
<w:style w:type="paragraph" w:styleId="DocNumber"><w:name w:val="Doc Number"/></w:style>
<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>
</w:styles>`;

const HEADER =
	`<w:hdr xmlns:w="${W}"><w:p>${textRun('No. ')}` +
	run('<w:fldChar w:fldCharType="begin"/>') +
	run('<w:instrText xml:space="preserve"> STYLEREF "Doc Number" \\* MERGEFORMAT </w:instrText>') +
	run('<w:fldChar w:fldCharType="separate"/>') +
	textRun('Number') +
	run('<w:fldChar w:fldCharType="end"/>') +
	'</w:p></w:hdr>';

/**
 * The whole package; `protect` adds forms protection, `body` replaces the form's body, `header`
 * the paragraphs of its header, and `styles` adds style definitions.
 */
export async function formTemplateFixture(
	options: { protect?: boolean; body?: string; header?: string; styles?: string } = {},
): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/settings" Target="settings.xml"/><Relationship Id="rId3" Type="${R}/header" Target="header1.xml"/></Relationships>`,
	);
	zip.file('word/styles.xml', STYLES.replace('</w:styles>', `${options.styles ?? ''}</w:styles>`));
	zip.file(
		'word/settings.xml',
		`<w:settings xmlns:w="${W}">${options.protect ? '<w:documentProtection w:edit="forms" w:enforcement="1"/>' : ''}</w:settings>`,
	);
	zip.file(
		'word/header1.xml',
		options.header ? `<w:hdr xmlns:w="${W}" xmlns:w14="${W14}">${options.header}</w:hdr>` : HEADER,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}"><w:body>${options.body ?? formTemplateBody()}` +
			'<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:pgSz w:w="12240" w:h="15840"/>' +
			'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}
