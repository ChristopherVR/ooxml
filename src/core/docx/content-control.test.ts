import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx } from './index';
import type { ContentControl, Paragraph, TextRun } from './model';
import { expectParagraph } from './test-support/access';
import { WORD_NS } from './xml';

const fixturePath = new URL(
	'./__fixtures__/content-controls/content-controls.docx',
	import.meta.url,
);
const load = async () => loadDocx(new Uint8Array(await readFile(fixturePath)));
const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;
const controlOf = (paragraph: Paragraph): ContentControl =>
	paragraph.runs.find((run) => run.contentControl)!.contentControl!;
const controlRuns = (paragraph: Paragraph): TextRun[] =>
	paragraph.runs.filter((run) => run.contentControl);

describe('content controls (w:sdt) authored by Word', () => {
	it('reads inline controls of every common type with their properties', async () => {
		const { model } = await load();
		const [name, agreed, colour, city, date] = model.blocks
			.slice(0, 5)
			.map((b) => expectParagraph(b));
		expect(controlOf(name!)).toMatchObject({
			level: 'inline',
			kind: 'text',
			alias: 'Full name',
			tag: 'fullName',
			lock: 'sdtLocked',
			showingPlaceholder: true,
		});
		expect(controlRuns(name!)[0]?.text).toBe('Click or tap here to enter text.');
		expect(controlOf(agreed!)).toMatchObject({
			kind: 'checkbox',
			checked: true,
			checkedGlyph: { code: '2612', font: 'MS Gothic' },
			uncheckedGlyph: { code: '2610', font: 'MS Gothic' },
		});
		expect(controlOf(colour!)).toMatchObject({
			kind: 'dropDownList',
			items: [
				{ display: 'Red', value: 'r' },
				{ display: 'Green', value: 'g' },
				{ display: 'Blue', value: 'b' },
			],
		});
		expect(controlRuns(colour!)[0]?.text).toBe('Green');
		expect(controlOf(city!)).toMatchObject({ kind: 'comboBox', showingPlaceholder: true });
		expect(controlOf(date!)).toMatchObject({
			kind: 'date',
			dateFormat: 'dd MMMM yyyy',
			dateLocale: 'en-US',
		});
		// Text before the control is ordinary text of the same paragraph.
		expect(name!.runs[0]?.contentControl).toBeUndefined();
		expect(name!.runs[0]?.text).toBe('Name: ');
	});

	it('reads a block-level control around a paragraph', async () => {
		const { model } = await load();
		const notes = expectParagraph(model.blocks[5]);
		expect(notes.runs[0]?.text).toBe('Notes');
		expect(notes.contentControl).toMatchObject({
			level: 'block',
			kind: 'richText',
			alias: 'Notes block',
			lock: 'contentLocked',
		});
		expect(expectParagraph(model.blocks[6]).contentControl).toBeUndefined();
	});

	it('keeps every control byte-for-byte when nothing changed', async () => {
		const loaded = await load();
		const before = await documentXml(new Uint8Array(await readFile(fixturePath)));
		const after = await documentXml(await loaded.save(loaded.model));
		const sdts = (xml: string) => xml.match(/<w:sdt>.*?<\/w:sdt>/gs);
		expect(sdts(after)).toEqual(sdts(before));
	});

	it('writes a toggled checkbox, a chosen list item and a typed value', async () => {
		const loaded = await load();
		const [, agreed, colour, city] = loaded.model.blocks.map((b) => expectParagraph(b));
		const box = controlRuns(agreed!)[0]!;
		box.text = '☐';
		box.contentControl = { ...box.contentControl!, checked: false };
		const blue = controlRuns(colour!)[0]!;
		blue.text = 'Blue';
		const cityRun = controlRuns(city!)[0]!;
		cityRun.text = 'Rome';
		cityRun.contentControl = { ...cityRun.contentControl!, showingPlaceholder: false };
		delete cityRun.style;
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w14:checked w14:val="0"/>');
		expect(xml).toContain('☐');
		expect(xml).toContain('<w:t>Blue</w:t>');
		const cityXml = /<w:sdt><w:sdtPr><w:alias w:val="City"\/>.*?<\/w:sdt>/s.exec(xml)![0];
		expect(cityXml).not.toContain('showingPlcHdr');
		expect(cityXml).toContain('<w:comboBox><w:listItem w:displayText="Paris"');
		const reopened = await loadDocx(await loaded.save(loaded.model));
		expect(controlOf(expectParagraph(reopened.model.blocks[1]))).toMatchObject({
			kind: 'checkbox',
			checked: false,
		});
		expect(controlRuns(expectParagraph(reopened.model.blocks[2]))[0]?.text).toBe('Blue');
	});

	it('writes edited properties and a selected date back into the same sdtPr', async () => {
		const loaded = await load();
		const [name, , , , date] = loaded.model.blocks.map((b) => expectParagraph(b));
		const run = controlRuns(name!)[0]!;
		run.contentControl = {
			...run.contentControl!,
			alias: 'Applicant',
			tag: 'who',
			lock: 'contentLocked',
		};
		const pick = controlRuns(date!)[0]!;
		pick.text = '05 March 2026';
		pick.contentControl = {
			...pick.contentControl!,
			showingPlaceholder: false,
			fullDate: '2026-03-05T00:00:00Z',
		};
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:alias w:val="Applicant"/><w:tag w:val="who"/>');
		expect(xml).toContain('<w:lock w:val="contentLocked"/>');
		expect(xml).toContain('w:fullDate="2026-03-05T00:00:00Z"');
		expect(xml).toContain('<w:t>05 March 2026</w:t>');
		expect(
			new RegExp(`<w:sdt>(?:(?!</w:sdt>).)*<w:dateFormat w:val="dd MMMM yyyy"/>`, 's').test(xml),
		).toBe(true);
	});

	it('edits the text of a block-level control and adds a paragraph inside it', async () => {
		const loaded = await load();
		const notes = expectParagraph(loaded.model.blocks[5]);
		notes.runs = [{ text: 'Changed' }];
		const extra: Paragraph = {
			type: 'paragraph',
			id: 'p-new',
			runs: [{ text: 'Second line' }],
			contentControl: notes.contentControl!,
		};
		loaded.model.blocks.splice(6, 0, extra);
		const xml = await documentXml(await loaded.save(loaded.model));
		const block = /<w:sdt><w:sdtPr><w:alias w:val="Notes block"\/>.*?<\/w:sdt>/s.exec(xml)![0];
		expect(block).toContain('<w:t>Changed</w:t>');
		expect(block).toContain('<w:t>Second line</w:t>');
		expect(block.indexOf('Changed')).toBeLessThan(block.indexOf('Second line'));
		expect(xml.indexOf('After')).toBeGreaterThan(xml.indexOf('Second line'));
		const reopened = await loadDocx(await loaded.save(loaded.model));
		expect(reopened.model.blocks.map((b) => b.contentControl?.alias)).toEqual([
			undefined,
			undefined,
			undefined,
			undefined,
			undefined,
			'Notes block',
			'Notes block',
			undefined,
		]);
	});
});

const hand = async (body: string) => {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
};

describe('content controls in tables and with unmodeled properties', () => {
	it('reads row and cell controls and keeps their markup', async () => {
		const loaded = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
				`<w:sdt><w:sdtPr><w:alias w:val="Row"/><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:tr>` +
				`<w:sdt><w:sdtPr><w:alias w:val="Cell"/><w:id w:val="8"/><w:text/></w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>` +
				`<w:tc><w:sdt><w:sdtPr><w:alias w:val="In cell"/><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>B</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc>` +
				`</w:tr></w:sdtContent></w:sdt></w:tbl>`,
		);
		const table = loaded.model.blocks[0]!;
		if (table.type !== 'table') throw new Error('table expected');
		expect(table.rowProperties?.[0]?.contentControl?.alias).toBe('Row');
		expect(table.rows[0]?.[0]?.contentControl).toMatchObject({
			alias: 'Cell',
			kind: 'text',
			level: 'cell',
		});
		expect(table.rows[0]?.[0]?.paragraphs[0]?.runs[0]?.text).toBe('A');
		expect(table.rows[0]?.[1]?.paragraphs[0]?.contentControl?.alias).toBe('In cell');
		table.rows[0]![0]!.paragraphs[0]!.runs = [{ text: 'A2' }];
		table.rows[0]![1]!.paragraphs[0]!.runs = [{ text: 'B2' }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:alias w:val="Row"/>');
		expect(xml).toContain('<w:alias w:val="Cell"/>');
		expect(xml).toContain('<w:alias w:val="In cell"/>');
		expect(xml).toContain('<w:t>A2</w:t>');
		expect(xml).toContain('<w:t>B2</w:t>');
	});

	it('keeps data binding and unknown properties verbatim', async () => {
		const loaded = await hand(
			`<w:p><w:sdt><w:sdtPr><w:alias w:val="Bound"/><w:id w:val="5"/><w:dataBinding w:prefixMappings="xmlns:ns0='urn:x'" w:xpath="/ns0:a/ns0:b" w:storeItemID="{11111111-2222-3333-4444-555555555555}"/><w:text/></w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent><w:r><w:t>value</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(controlOf(paragraph).dataBound).toBe(true);
		paragraph.runs[0]!.text = 'changed';
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('w:xpath="/ns0:a/ns0:b"');
		expect(xml).toContain('<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>');
		expect(xml).toContain('<w:t>changed</w:t>');
	});

	it('keeps the end properties of an inline control it writes anew', async () => {
		const loaded = await hand(
			`<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="57"/><w:text/></w:sdtPr>` +
				`<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent><w:r><w:t>A. Writer</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		// A run before the control moves it to a run index that has no source control there.
		paragraph.runs.splice(1, 0, { text: '(author) ', italic: true });
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toMatch(
			/<w:sdtPr[^>]*><w:id w:val="57"\/><w:text\/><\/w:sdtPr><w:sdtEndPr[^>]*><w:rPr><w:b\/><\/w:rPr><\/w:sdtEndPr><w:sdtContent>/,
		);
	});

	it('creates a new inline control from the model', async () => {
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs.push({
			text: 'Red',
			contentControl: {
				id: '4242',
				level: 'inline',
				kind: 'dropDownList',
				alias: 'Colour',
				items: [
					{ display: 'Red', value: 'r' },
					{ display: 'Blue', value: 'b' },
				],
			},
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(
			'<w:sdt><w:sdtPr><w:alias w:val="Colour"/><w:id w:val="4242"/><w:dropDownList><w:listItem w:displayText="Red" w:value="r"/>',
		);
		const reopened = await loadDocx(await loaded.save(loaded.model));
		expect(controlOf(expectParagraph(reopened.model.blocks[0])).items).toHaveLength(2);
	});

	it('writes verbatim properties of a new control only as its properties', async () => {
		// A control pasted from another page carries what that page wrote: anything but w:sdtPr and
		// w:sdtEndPr is not written, here fields that would fetch a picture or text when updated.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const field = (code: string) =>
			`<w:fldSimple w:instr="${code}"><w:r><w:t>x</w:t></w:r></w:fldSimple>`;
		paragraph.runs.push({
			text: 'Red',
			contentControl: {
				id: '4243',
				level: 'inline',
				kind: 'richText',
				sourceXml: `<w:p xmlns:w="${WORD_NS}">${field('INCLUDEPICTURE &quot;http://example.test/a.png&quot;')}</w:p>`,
				endSourceXml: `<w:sdtContent xmlns:w="${WORD_NS}"><w:r>${field('INCLUDETEXT &quot;http://example.test/b&quot;')}</w:r></w:sdtContent>`,
			},
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('INCLUDE');
		expect(xml).toContain('<w:sdt><w:sdtPr><w:id w:val="4243"/></w:sdtPr><w:sdtContent>');
		// Properties that are what they say are kept as they were.
		paragraph.runs[1]!.contentControl = {
			id: '4243',
			level: 'inline',
			kind: 'richText',
			alias: 'Kept',
			sourceXml: `<w:sdtPr xmlns:w="${WORD_NS}"><w:alias w:val="Kept"/><w:id w:val="1"/></w:sdtPr>`,
			endSourceXml: `<w:sdtEndPr xmlns:w="${WORD_NS}"><w:rPr><w:b/></w:rPr></w:sdtEndPr>`,
		};
		expect(await documentXml(await loaded.save(loaded.model))).toMatch(
			/<w:sdtPr[^>]*><w:alias w:val="Kept"\/><w:id w:val="4243"\/><\/w:sdtPr><w:sdtEndPr[^>]*><w:rPr><w:b\/><\/w:rPr><\/w:sdtEndPr>/,
		);
	});
});
