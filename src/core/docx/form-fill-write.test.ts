import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { listFormFields } from './form-fill-fields';
import { fillFormFields } from './form-fill-write';
import {
	formTemplateBody,
	formTemplateFixture,
	paragraph,
	textRun,
} from './test-support/form-template-fixture';

async function parts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const zip = await JSZip.loadAsync(bytes);
	const files = new Map<string, Uint8Array>();
	for (const entry of Object.values(zip.files))
		if (!entry.dir) files.set(entry.name, await entry.async('uint8array'));
	return files;
}
const part = async (bytes: Uint8Array, name: string) =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';

async function fill(values: Record<string, string | boolean | number>, protect = false) {
	const original = await formTemplateFixture({ protect });
	const loaded = await loadDocx(original);
	const { model, results } = fillFormFields(loaded.model, values);
	const saved = await loaded.save(model);
	return { original, saved, results, reloaded: listFormFields((await loadDocx(saved)).model) };
}
const valueOf = (list: ReturnType<typeof listFormFields>, id: string) =>
	list.find((field) => field.id === id)?.value;

describe('filling a form', () => {
	it('sets each kind of field and saves it as Word would', async () => {
		const { saved, results, reloaded } = await fill({
			'field#1': 'North wing',
			'field:IssueDate': '2026-10-21',
			'field#3': 'DOC-0042',
			'field#4': 'b',
			'field:Status': 'Revised',
			'field:Approved': true,
			'field:Count': 1234.5,
			'control:101': true,
			'control:201': 'Reference',
			'control:301': 'Example client',
			'control:302': '2026-11-02',
			'control:303': 'Elsewhere',
		});
		expect(results.every((result) => result.ok)).toBe(true);
		expect(results.find((result) => result.id === 'field:IssueDate')?.value).toBe('21-Oct-2026');
		expect(valueOf(reloaded, 'field#1')).toBe('North wing');
		expect(valueOf(reloaded, 'field:IssueDate')).toBe('21-Oct-2026');
		expect(valueOf(reloaded, 'field#3')).toBe('DOC-0042');
		expect(valueOf(reloaded, 'field#4')).toBe('B');
		expect(valueOf(reloaded, 'field:Status')).toBe('Revised');
		expect(valueOf(reloaded, 'field:Approved')).toBe(true);
		expect(valueOf(reloaded, 'field:Count')).toBe('1,234.50');
		expect(valueOf(reloaded, 'control:101')).toBe(true);
		expect(valueOf(reloaded, 'control:301')).toBe('Example client');
		expect(valueOf(reloaded, 'control:302')).toBe('02 November 2026');
		expect(valueOf(reloaded, 'control:303')).toBe('Elsewhere');
		// A temporary control is gone once it is filled, as in Word; its text stays.
		expect(reloaded.some((field) => field.id === 'control:201')).toBe(false);
		const xml = await part(saved, 'word/document.xml');
		expect(xml).toContain('>Reference<');
		expect(xml).not.toMatch(/<w:id w:val="201"\/>/);
		expect(xml).toContain('<w:id w:val="202"/>');
		// The client control no longer shows its placeholder, in text or style.
		const client = /<w:sdt><w:sdtPr><w:alias w:val="Client"\/>.*?<\/w:sdt>/s.exec(xml)?.[0] ?? '';
		expect(client).not.toContain('showingPlcHdr');
		expect(client).not.toContain('PlaceholderText');
		// The header repeats the document number through STYLEREF.
		expect(await part(saved, 'word/header1.xml')).toContain('>DOC-0042<');
	});

	it('refuses values that do not fit their field and says why', async () => {
		const { results, saved, original } = await fill({
			'field:Status': 'Closed',
			'field#3': 'DOC-0042-ABCDE',
			'field:IssueDate': '21/10/2026',
			'field:Count': 'twelve',
			'field:Approved': 'yes',
			'field:Archived': 'Yes',
			'control:304': 'Somebody',
			'field#99': 'x',
			'field#1': 'two\nlines',
		});
		expect(Object.fromEntries(results.map((result) => [result.id, result.reason]))).toEqual({
			'field:Status': 'not one of the entries: New, Revised, Withdrawn',
			'field#3': "longer than the field's 12 characters",
			'field:IssueDate': 'expected a date as YYYY-MM-DD',
			'field:Count': 'expected a number',
			'field:Approved': 'expected true or false',
			'field:Archived': 'disabled',
			'control:304': 'locked',
			'field#99': 'no field has this id',
			'field#1': 'the field takes one line',
		});
		expect(results.every((result) => !result.ok)).toBe(true);
		expect(saved).toEqual(original);
	});

	it('leaves every part it did not fill as it was, byte for byte', async () => {
		const { original, saved } = await fill({ 'field#18': 'Initial issue' });
		const before = await parts(original);
		const after = await parts(saved);
		expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
		for (const [name, bytes] of before)
			if (name !== 'word/document.xml') expect(after.get(name), name).toEqual(bytes);
		// In the body, only the filled paragraph changed.
		const tables = (xml: string) => xml.match(/<w:tbl>.*?<\/w:tbl>/gs) ?? [];
		const was = await part(original, 'word/document.xml');
		const now = await part(saved, 'word/document.xml');
		expect(tables(now)).toEqual(tables(was));
		expect(now).toContain('>Initial issue<');
	});

	it('keeps the form fillable: field data, names and bookmarks stay', async () => {
		const { original, saved } = await fill({
			'field:IssueDate': '2026-10-21',
			'field#5': 'Team A',
			'field:Approved': true,
		});
		const before = listFormFields((await loadDocx(original)).model);
		const after = listFormFields((await loadDocx(saved)).model);
		expect(after.map((field) => [field.id, field.type, field.fillable])).toEqual(
			before.map((field) => [field.id, field.type, field.fillable]),
		);
		const xml = await part(saved, 'word/document.xml');
		expect(xml.match(/<w:ffData>/g)).toHaveLength(18);
		expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="IssueDate"\/>/);
		expect(xml).toContain('<w:format w:val="dd-MMM-yyyy"/>');
	});

	it('clears a field back to its placeholder', async () => {
		const loaded = await loadDocx(await formTemplateFixture());
		const filled = fillFormFields(loaded.model, { 'control:301': 'Acme', 'field#1': 'East' });
		const cleared = fillFormFields(filled.model, { 'control:301': '', 'field#1': '' });
		expect(cleared.results.every((result) => result.ok)).toBe(true);
		const xml = await part(await loaded.save(cleared.model), 'word/document.xml');
		const client = /<w:sdt><w:sdtPr><w:alias w:val="Client"\/>.*?<\/w:sdt>/s.exec(xml)?.[0] ?? '';
		expect(client).toContain('<w:showingPlcHdr/>');
		expect(client).toContain('<w:rStyle w:val="PlaceholderText"/>');
		expect(client).toContain('>Click or tap here to enter text.<');
		const fields = listFormFields(cleared.model);
		expect(fields.find((field) => field.id === 'field#1')?.value).toBe('');
	});

	it('changes nothing for a value a field already holds', async () => {
		const { original, saved, results } = await fill({
			'field#1': 'Project',
			'field:Approved': false,
			'control:303': 'North',
		});
		expect(results.every((result) => result.ok)).toBe(true);
		expect(saved).toEqual(original);
	});

	it('fills a document protected for forms and keeps its protection', async () => {
		const { original, saved, reloaded } = await fill({ 'field#9': 'R. Doe' }, true);
		expect(valueOf(reloaded, 'field#9')).toBe('R. Doe');
		expect(await part(saved, 'word/settings.xml')).toBe(await part(original, 'word/settings.xml'));
	});

	it('keeps how a result is split into runs when a run holds properties it does not model', async () => {
		const result = [textRun(' '), textRun(' ', '<w:effect w:val="sparkle"/>'), textRun(' ')].join(
			'',
		);
		const body = paragraph(
			textRun('Lot: ') +
				'<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val=""/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
				'<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
				`${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
		);
		const loaded = await loadDocx(await formTemplateFixture({ body }));
		const { model, results } = fillFormFields(loaded.model, { 'field#1': 'L-77' });
		expect(results).toEqual([{ id: 'field#1', ok: true, value: 'L-77' }]);
		const saved = await loaded.save(model);
		expect(listFormFields((await loadDocx(saved)).model)[0]?.value).toBe('L-77');
		const xml = await part(saved, 'word/document.xml');
		expect(xml).toContain('<w:effect w:val="sparkle"/>');
		// The result keeps its three runs: the first holds the text, the others are empty.
		const between = xml.split('w:fldCharType="separate"/>')[1]!.split('w:fldCharType="end"')[0]!;
		expect(between.match(/<w:r>|<w:r /g)).toHaveLength(4);
		expect(between).toContain('>L-77<');
	});

	it('fills a control around paragraphs, keeping one paragraph with line breaks', async () => {
		const body =
			paragraph(textRun('Summary')) +
			'<w:sdt><w:sdtPr><w:id w:val="401"/><w:showingPlcHdr/><w:text w:multiLine="1"/></w:sdtPr><w:sdtContent>' +
			paragraph(
				textRun('Click or tap here to enter text.', '<w:rStyle w:val="PlaceholderText"/>'),
			) +
			paragraph(textRun('')) +
			'</w:sdtContent></w:sdt>' +
			paragraph(textRun('After'));
		const loaded = await loadDocx(await formTemplateFixture({ body }));
		expect(listFormFields(loaded.model)).toMatchObject([
			{ id: 'control:401', label: 'Summary', type: 'text', multiLine: true, value: '' },
		]);
		const { model, results } = fillFormFields(loaded.model, {
			'control:401': 'Line one\nLine two',
		});
		expect(results).toEqual([{ id: 'control:401', ok: true, value: 'Line one\nLine two' }]);
		const saved = await loaded.save(model);
		const xml = await part(saved, 'word/document.xml');
		const control = /<w:sdt>.*?<\/w:sdt>/s.exec(xml)?.[0] ?? '';
		expect(control).toMatch(/>Line one<\/w:t><w:br\/><w:t>Line two</);
		expect(control).not.toContain('showingPlcHdr');
		expect(control.match(/<w:p>|<w:p /g)).toHaveLength(1);
		expect(listFormFields((await loadDocx(saved)).model)[0]?.value).toBe('Line one\nLine two');
	});

	it('leaves a legacy field inside a control that locks its content as it is', async () => {
		const body = paragraph(
			textRun('Ref: ') +
				'<w:sdt><w:sdtPr><w:id w:val="501"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
				'<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Ref"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
				'<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
				`${textRun('R-1')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:sdtContent></w:sdt>`,
		);
		const loaded = await loadDocx(await formTemplateFixture({ body }));
		expect(listFormFields(loaded.model)).toMatchObject([
			{ id: 'field:Ref', label: 'Ref', value: 'R-1', fillable: false, reason: 'locked' },
		]);
		expect(fillFormFields(loaded.model, { 'field:Ref': 'R-2' }).results).toEqual([
			{ id: 'field:Ref', ok: false, reason: 'locked' },
		]);
	});

	it('fills a field in a header and leaves the body as it was', async () => {
		const header = paragraph(
			textRun('Version: ') +
				'<w:sdt><w:sdtPr><w:tag w:val="version"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>' +
				`${textRun('Click or tap here to enter text.', '<w:rStyle w:val="PlaceholderText"/>')}</w:sdtContent></w:sdt>`,
		);
		const original = await formTemplateFixture({ header });
		const loaded = await loadDocx(original);
		const field = listFormFields(loaded.model).at(-1)!;
		expect(field).toMatchObject({ id: 'tag:version', label: 'Version', story: 'header' });
		const saved = await loaded.save(fillFormFields(loaded.model, { 'tag:version': '3.1' }).model);
		expect(await part(saved, 'word/header1.xml')).toContain('>3.1<');
		// The section's own properties are the section writer's concern: the body's blocks stay.
		const blocks = async (bytes: Uint8Array) =>
			(await part(bytes, 'word/document.xml')).split('<w:sectPr')[0];
		expect(await blocks(saved)).toBe(await blocks(original));
	});

	it('brings body STYLEREF results up to date unless asked not to', async () => {
		const reference =
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> STYLEREF "Doc Number" </w:instrText></w:r>' +
			`<w:r><w:fldChar w:fldCharType="separate"/></w:r>${textRun('Number')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
		const body = formTemplateBody() + paragraph(textRun('See ') + reference);
		const loaded = await loadDocx(await formTemplateFixture({ body }));
		const text = (model: typeof loaded.model) => {
			const last = model.blocks.at(-1)!;
			return last.type === 'paragraph' ? last.runs.map((run) => run.text).join('') : '';
		};
		expect(text(fillFormFields(loaded.model, { 'field#3': 'DOC-7' }).model)).toBe('See DOC-7');
		expect(
			text(fillFormFields(loaded.model, { 'field#3': 'DOC-7' }, { updateStyleRefs: false }).model),
		).toBe('See Number');
	});
});
