import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import {
	formFieldDisplayText,
	formFieldOfRun,
	selectFormDropDownEntry,
	setFormCheckbox,
} from './form-fields';
import { at, expectParagraph, expectTable, must } from './test-support/access';
import type { Paragraph, TextRun } from './model';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (inner: string, props = '') =>
	`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}${inner}</w:r>`;
const text = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`, '<w:noProof/>');

/** A legacy form field as Word writes it: bookmark, begin with ffData, code, result, end. */
function formField(name: string, ffData: string, code: string, result: string | null, id = 1) {
	return (
		`<w:bookmarkStart w:id="${id}" w:name="${name}"/>` +
		run(
			`<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/><w:enabled/><w:calcOnExit w:val="0"/>${ffData}</w:ffData></w:fldChar>`,
		) +
		run(`<w:instrText xml:space="preserve"> ${code} </w:instrText>`) +
		(result === null ? '' : run('<w:fldChar w:fldCharType="separate"/>') + result) +
		run('<w:fldChar w:fldCharType="end"/>') +
		`<w:bookmarkEnd w:id="${id}"/>`
	);
}

const ENSPACES = '\u2002'.repeat(5);
const textField = formField(
	'IssueDate',
	'<w:helpText w:type="text" w:val="Enter the issue date"/><w:statusText w:type="text" w:val="Issue date"/>' +
		'<w:textInput><w:type w:val="date"/><w:maxLength w:val="11"/><w:format w:val="dd-MMM-yyyy"/></w:textInput>',
	'FORMTEXT',
	Array.from(ENSPACES, (space) => text(space)).join(''),
);
const checkboxField = formField(
	'Approved',
	'<w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>',
	'FORMCHECKBOX',
	null,
	2,
);
const dropDownField = formField(
	'Status',
	'<w:ddList><w:result w:val="1"/><w:listEntry w:val="New"/><w:listEntry w:val="Revised"/><w:listEntry w:val="Withdrawn"/></w:ddList>',
	'FORMDROPDOWN',
	null,
	3,
);

async function load(body: string, settings?: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	if (settings)
		zip.file('word/settings.xml', `<w:settings xmlns:w="${W}">${settings}</w:settings>`);
	const original = await zip.generateAsync({ type: 'uint8array' });
	return { loaded: await loadDocx(original), original };
}

async function documentXml(bytes: Uint8Array): Promise<string> {
	return (await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';
}

const beginOf = (paragraph: Paragraph): TextRun =>
	must(paragraph.runs.find((item) => item.formField));

describe('legacy form fields', () => {
	it('reads a text field: name, help, status, type, length, format and its blank result', async () => {
		const { loaded } = await load(`<w:p>${textField}</w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(beginOf(paragraph).formField).toMatchObject({
			kind: 'text',
			name: 'IssueDate',
			enabled: true,
			calcOnExit: false,
			helpText: { type: 'text', value: 'Enter the issue date' },
			statusText: { type: 'text', value: 'Issue date' },
			textInput: { type: 'date', maxLength: 11, format: 'dd-MMM-yyyy' },
		});
		expect(formFieldDisplayText(beginOf(paragraph))).toBeUndefined();
		const result = paragraph.runs.filter((item) => item.field?.instr === 'FORMTEXT');
		expect(result.map((item) => item.text).join('')).toBe(ENSPACES);
	});

	it('reads a check box and a drop-down, which have no result text, and shows their values', async () => {
		const { loaded } = await load(
			`<w:p>${checkboxField}${run('<w:t xml:space="preserve"> Yes </w:t>')}${dropDownField}</w:p>`,
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const [checkbox, dropDown] = paragraph.runs.filter((item) => item.formField);
		expect(must(checkbox).formField).toMatchObject({
			kind: 'checkbox',
			name: 'Approved',
			checkBox: { default: false, sizeAuto: true },
		});
		expect(must(dropDown).formField).toMatchObject({
			kind: 'dropDown',
			name: 'Status',
			dropDown: { entries: ['New', 'Revised', 'Withdrawn'], result: 1 },
		});
		expect(formFieldDisplayText(must(checkbox))).toBe('☐');
		expect(formFieldDisplayText(must(dropDown))).toBe('Revised');
	});

	it('fills a text field in a body paragraph and keeps its data and bookmark', async () => {
		const { loaded } = await load(`<w:p>${text('Issued: ')}${textField}</w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const results = paragraph.runs.filter((item) => item.field?.instr === 'FORMTEXT');
		at(results, 0).text = '20-Oct-2026';
		for (const extra of results.slice(1)) extra.text = '';
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain(
			'<w:ffData><w:name w:val="IssueDate"/><w:enabled/><w:calcOnExit w:val="0"/><w:helpText w:type="text" w:val="Enter the issue date"/>',
		);
		const field = xml.slice(xml.indexOf('<w:bookmarkStart'), xml.indexOf('<w:bookmarkEnd'));
		expect(field).toContain('w:fldCharType="begin"');
		expect(field).toContain('>20-Oct-2026<');
		expect(field).toContain('w:fldCharType="end"');
		expect(field).not.toContain('\u2002');
	});

	it('toggles a check box by writing w:checked and keeps the rest of the field data', async () => {
		const { loaded } = await load(
			`<w:p>${checkboxField.replace('<w:enabled/>', '<w:enabled/><w:tabIndex w:val="4"/><w:entryMacro w:val="Check"/>')}</w:p>`,
		);
		const begin = beginOf(expectParagraph(loaded.model.blocks[0]));
		begin.formField = setFormCheckbox(must(begin.formField), true);
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain(
			'<w:ffData><w:name w:val="Approved"/><w:enabled/><w:tabIndex w:val="4"/><w:entryMacro w:val="Check"/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked/></w:checkBox></w:ffData>',
		);
		const reread = beginOf(expectParagraph((await loadDocx(await loaded.save())).model.blocks[0]));
		expect(reread.formField?.checkBox?.checked).toBe(true);
		expect(formFieldDisplayText(reread)).toBe('☒');
	});

	it('fills a field whose data is indented inside its begin marker', async () => {
		const indented = checkboxField
			.replace('<w:ffData>', '\n\t<w:ffData>\n\t\t')
			.replace('</w:ffData>', '\n\t</w:ffData>\n');
		const { loaded } = await load(`<w:p>${indented}</w:p>`);
		const begin = beginOf(expectParagraph(loaded.model.blocks[0]));
		begin.formField = setFormCheckbox(must(begin.formField), true);
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w:checked/>');
		expect(xml).toContain('<w:name w:val="Approved"/>');
	});

	it('picks a drop-down entry by writing w:result', async () => {
		const { loaded } = await load(`<w:p>${dropDownField}</w:p>`);
		const begin = beginOf(expectParagraph(loaded.model.blocks[0]));
		begin.formField = must(selectFormDropDownEntry(must(begin.formField), 2));
		expect(selectFormDropDownEntry(must(begin.formField), 3)).toBeUndefined();
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain(
			'<w:ddList><w:result w:val="2"/><w:listEntry w:val="New"/><w:listEntry w:val="Revised"/><w:listEntry w:val="Withdrawn"/></w:ddList>',
		);
	});

	it('fills fields in table cells', async () => {
		const cell = (content: string) => `<w:tc><w:p>${content}</w:p></w:tc>`;
		const { loaded } = await load(
			`<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
				`<w:tr>${cell(text('Date'))}${cell(textField)}</w:tr>` +
				`<w:tr>${cell(text('Status'))}${cell(dropDownField)}</w:tr></w:tbl>`,
		);
		const table = expectTable(loaded.model.blocks[0]);
		const date = at(at(at(table.rows, 0), 1).paragraphs, 0);
		const results = date.runs.filter((item) => item.field?.instr === 'FORMTEXT');
		at(results, 0).text = '01-Jan-2027';
		for (const extra of results.slice(1)) extra.text = '';
		const status = beginOf(at(at(at(table.rows, 1), 1).paragraphs, 0));
		status.formField = must(selectFormDropDownEntry(must(status.formField), 0));
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('>01-Jan-2027<');
		expect(xml).toContain('<w:result w:val="0"/>');
	});

	describe('the bookmark Word gives each field its name', () => {
		const plain = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);
		const rejected = checkboxField
			.replaceAll('Approved', 'Rejected')
			.replaceAll('w:id="2"', 'w:id="4"');
		/** For each bookmark in `xml`, the names of the form fields it holds. */
		const bookmarked = (xml: string) =>
			Object.fromEntries(
				[
					...xml.matchAll(
						/<w:bookmarkStart w:id="(\d+)" w:name="(\w+)"\/>([\s\S]*?)<w:bookmarkEnd w:id="\1"\/>/g,
					),
				].map(([, , name, inner]) => [
					name,
					[...inner!.matchAll(/<w:name w:val="(\w+)"\/>/g)].map((found) => found[1]),
				]),
			);
		/** A paragraph of two check boxes with labels: begin, code and end runs for each field. */
		async function twoBoxes() {
			const { loaded } = await load(
				`<w:p>${checkboxField}${plain(' Yes ')}${rejected}${plain(' No')}</w:p>`,
			);
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const runs = paragraph.runs;
			expect(runs).toHaveLength(8);
			return {
				loaded,
				paragraph,
				approved: runs.slice(0, 3),
				yes: at(runs, 3),
				rejected: runs.slice(4, 7),
				no: at(runs, 7),
			};
		}

		it('keeps each field in its own bookmark when two fields swap places', async () => {
			const { loaded, paragraph, approved, yes, rejected, no } = await twoBoxes();
			paragraph.runs = [...rejected, yes, ...approved, no];
			const xml = await documentXml(await loaded.save());
			expect(bookmarked(xml)).toEqual({ Rejected: ['Rejected'], Approved: ['Approved'] });
		});

		it('takes a field its bookmark along when the field moves in its paragraph', async () => {
			const { loaded, paragraph, approved, yes, rejected, no } = await twoBoxes();
			paragraph.runs = [yes, ...rejected, no, ...approved];
			const xml = await documentXml(await loaded.save());
			expect(bookmarked(xml)).toEqual({ Rejected: ['Rejected'], Approved: ['Approved'] });
			expect(xml.indexOf('w:name="Approved"')).toBeGreaterThan(xml.indexOf('> No<'));
		});
	});

	it('leaves the package untouched when nothing changes', async () => {
		const { loaded, original } = await load(
			`<w:p>${textField}${checkboxField}${dropDownField}</w:p>`,
		);
		expect(await loaded.save()).toEqual(original);
	});

	it('finds the form field of a result run and of its begin marker', async () => {
		const { loaded } = await load(`<w:p>${textField}</w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const resultIndex = paragraph.runs.findIndex((item) => item.field?.instr === 'FORMTEXT');
		expect(formFieldOfRun(paragraph.runs, resultIndex)?.name).toBe('IssueDate');
		expect(formFieldOfRun(paragraph.runs, 0)?.name).toBe('IssueDate');
	});
});

describe('document protection and form field shading settings', () => {
	it('reads forms protection and the no-shading setting', async () => {
		const { loaded } = await load(
			`<w:p>${textField}</w:p>`,
			'<w:documentProtection w:edit="forms" w:enforcement="1"/><w:doNotShadeFormData/>',
		);
		expect(loaded.model.protection).toEqual({ edit: 'forms', enforced: true });
		expect(loaded.model.formFieldShading).toBe(false);
	});

	it('reads the sections a document protected for forms leaves unprotected', async () => {
		const { loaded } = await load(
			`<w:p><w:pPr><w:sectPr><w:formProt w:val="0"/></w:sectPr></w:pPr>${text('Notes')}</w:p>` +
				`<w:p>${textField}</w:p>`,
			'<w:documentProtection w:edit="forms" w:enforcement="1"/>',
		);
		expect(loaded.model.sections?.map((section) => section.formProtection)).toEqual([
			false,
			undefined,
		]);
		expect(loaded.model.warnings.join(' ')).toContain('sections left unprotected');
		at(expectParagraph(loaded.model.blocks[0]).runs, 0).text = 'More notes';
		expect(await documentXml(await loaded.save())).toContain('<w:formProt w:val="0"/>');
	});

	it('warns about restrictions it keeps but does not enforce, and not about none', async () => {
		const warnings = async (edit: string) =>
			(
				await load(
					`<w:p>${text('A')}</w:p>`,
					`<w:documentProtection w:edit="${edit}" w:enforcement="1"/>`,
				)
			).loaded.model.warnings.join(' ');
		expect(await warnings('readOnly')).toContain('limits editing to reading');
		expect(await warnings('none')).not.toMatch(/protect|limits editing/);
	});

	it('enforces protection whose enforcement is not given, as Word does', async () => {
		const { loaded } = await load(
			`<w:p>${textField}</w:p>`,
			'<w:documentProtection w:edit="forms"/>',
		);
		expect(loaded.model.protection).toEqual({ edit: 'forms', enforced: true });
		expect(loaded.model.formFieldShading).toBeUndefined();
	});

	it('treats protection whose enforcement is off as off', async () => {
		for (const off of ['0', 'false', 'off']) {
			const { loaded } = await load(
				`<w:p>${textField}</w:p>`,
				`<w:documentProtection w:edit="forms" w:enforcement="${off}"/>`,
			);
			expect(loaded.model.protection).toEqual({ edit: 'forms', enforced: false });
		}
	});
});
