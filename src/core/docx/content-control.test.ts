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

	it('writes the id and the kind the model gives a control around a paragraph', async () => {
		// A model edit can give a block another control than its w:sdt holds: patching that w:sdt
		// writes the new id, and the old kind's element goes, so the properties stay one control.
		const loaded = await hand(
			`<w:sdt><w:sdtPr><w:alias w:val="Pick"/><w:id w:val="5"/><w:dropDownList>` +
				`<w:listItem w:displayText="Yes" w:value="y"/></w:dropDownList></w:sdtPr>` +
				`<w:sdtContent><w:p><w:r><w:t>Yes</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.contentControl = {
			...paragraph.contentControl!,
			id: '77',
			kind: 'checkbox',
			checked: true,
		};
		paragraph.runs = [{ text: '\u2612' }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('w:dropDownList');
		expect(xml).not.toContain('<w:id w:val="5"/>');
		expect(xml.match(/<w14:checkbox>/g)).toHaveLength(1);
		const reopened = expectParagraph(
			(await loadDocx(await loaded.save(loaded.model))).model.blocks[0],
		);
		expect(reopened.contentControl).toMatchObject({ id: '77', kind: 'checkbox', checked: true });
	});

	it('replaces any type element when the model gives a control another kind', async () => {
		// w:richText may be written for the default type, and an entity picker is a type of Word's
		// extensions: either goes when the kind changes, so the properties hold one type.
		for (const type of ['<w:richText/>', '<w14:entityPicker/>']) {
			const loaded = await hand(
				`<w:sdt><w:sdtPr><w:id w:val="6"/>${type}</w:sdtPr>` +
					`<w:sdtContent><w:p><w:r><w:t>Yes</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
			);
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			paragraph.contentControl = { ...paragraph.contentControl!, kind: 'checkbox', checked: true };
			paragraph.runs = [{ text: '\u2612' }];
			const xml = await documentXml(await loaded.save(loaded.model));
			expect(xml, type).not.toContain(type.slice(1, -2));
			expect(xml.match(/<w14:checkbox>/g), type).toHaveLength(1);
		}
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

	it('leaves out the id a copy was copied with when the model gives it one Word cannot read', async () => {
		const loaded = await hand(
			`<w:p><w:sdt><w:sdtPr><w:alias w:val="Name"/><w:id w:val="21"/><w:text/></w:sdtPr>` +
				`<w:sdtContent><w:r><w:t>Ann</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const own = controlOf(paragraph);
		paragraph.runs.push({ text: 'Bo', contentControl: { ...own, id: 'abc' } });
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml.match(/<w:id w:val="21"\/>/g)).toHaveLength(1);
		expect(xml.match(/<w:alias w:val="Name"\/>/g)).toHaveLength(2);
	});

	it('writes only the list entries of a new control that have text', async () => {
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		expectParagraph(loaded.model.blocks[0]).runs.push({
			text: 'A',
			contentControl: {
				id: '4245',
				level: 'inline',
				kind: 'dropDownList',
				items: [null, { display: null, value: null }, 'x', { display: 'A', value: 'a' }] as never,
			},
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(
			'<w:dropDownList><w:listItem w:displayText="A" w:value="a"/></w:dropDownList>',
		);
	});

	it('writes the end properties of a new control only when they are run properties', async () => {
		// w:sdtEndPr holds the formatting of the control's end, run properties only: one that holds
		// anything else, such as a field one level down, is not written.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const field = `<w:fldSimple w:instr="INCLUDETEXT &quot;http://example.test/b&quot;"><w:r><w:t>x</w:t></w:r></w:fldSimple>`;
		const end = (inner: string) => `<w:sdtEndPr xmlns:w="${WORD_NS}">${inner}</w:sdtEndPr>`;
		paragraph.runs.push({
			text: 'Red',
			contentControl: { id: '4244', level: 'inline', kind: 'richText' },
		});
		for (const endSourceXml of [
			end(`<w:rPr><w:b/>${field}</w:rPr>`),
			end(`<w:rPr><w:b/></w:rPr><w:p>${field}</w:p>`),
			end(`<w:rPr><w:b/></w:rPr>text`),
		]) {
			paragraph.runs[1]!.contentControl = {
				id: '4244',
				level: 'inline',
				kind: 'richText',
				endSourceXml,
			};
			const xml = await documentXml(await loaded.save(loaded.model));
			expect(xml, endSourceXml).not.toContain('INCLUDETEXT');
			expect(xml, endSourceXml).toContain(
				'<w:sdt><w:sdtPr><w:id w:val="4244"/></w:sdtPr><w:sdtContent>',
			);
		}
		paragraph.runs[1]!.contentControl = {
			id: '4244',
			level: 'inline',
			kind: 'richText',
			endSourceXml: end('<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>'),
		};
		expect(await documentXml(await loaded.save(loaded.model))).toContain(
			'<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:sdtEndPr><w:sdtContent>',
		);
	});

	it('keeps verbatim properties that hold the attributes later versions of Word write', async () => {
		// Word 2021 and later put a checksum on a data binding (w16sdtdh) and a UTC date on a
		// formatting revision (w16du): a copy of such a control keeps its properties as Word wrote them.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const DH = 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash';
		const DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
		const binding = `<w:dataBinding w:xpath="/a" w:storeItemID="{1}" w16sdtdh:storeItemChecksum="Ab12"/>`;
		const revision =
			'<w:rPrChange w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z" w16du:dateUtc="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>';
		paragraph.runs.push({
			text: 'Red',
			contentControl: {
				id: '4245',
				level: 'inline',
				kind: 'text',
				sourceXml:
					`<w:sdtPr xmlns:w="${WORD_NS}" xmlns:w16sdtdh="${DH}" xmlns:w16du="${DU}">` +
					`<w:rPr><w:b/>${revision}</w:rPr><w:id w:val="4245"/>${binding}<w:text/></w:sdtPr>`,
				endSourceXml: `<w:sdtEndPr xmlns:w="${WORD_NS}" xmlns:w16du="${DU}"><w:rPr>${revision}</w:rPr></w:sdtEndPr>`,
			},
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('w16sdtdh:storeItemChecksum="Ab12"');
		expect(xml.match(/w16du:dateUtc=/g)).toHaveLength(2);
	});

	it('writes no verbatim properties with a second type element, one of an extension included', async () => {
		// w14:checkbox and a repeating section's elements are types too: properties with one of them
		// beside another type, or with an extension element twice, are built from the model.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
		const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
		const sdtPr = (inner: string) =>
			`<w:sdtPr xmlns:w="${WORD_NS}" xmlns:w14="${W14}" xmlns:w15="${W15}"><w:id w:val="4246"/>${inner}</w:sdtPr>`;
		paragraph.runs.push({
			text: 'Red',
			contentControl: { id: '4246', level: 'inline', kind: 'text' },
		});
		for (const sourceXml of [
			sdtPr('<w:text/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>'),
			sdtPr('<w:text/><w15:repeatingSection/>'),
			sdtPr('<w:text/><w15:appearance w15:val="hidden"/><w15:appearance w15:val="tags"/>'),
		]) {
			paragraph.runs[1]!.contentControl = { id: '4246', level: 'inline', kind: 'text', sourceXml };
			const xml = await documentXml(await loaded.save(loaded.model));
			expect(xml, sourceXml).toContain(
				'<w:sdt><w:sdtPr><w:id w:val="4246"/><w:text/></w:sdtPr><w:sdtContent>',
			);
		}
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

	it('writes the verbatim properties of a new control only when they are properties', async () => {
		// A control pasted from another page brings what that page wrote in its data-sdt attribute:
		// anything but a w:sdtPr there is not written, here a field that would fetch a picture.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const field = `<w:fldSimple w:instr="INCLUDEPICTURE &quot;http://example.test/a.png&quot;"><w:r><w:t>x</w:t></w:r></w:fldSimple>`;
		const control = (sourceXml: string): ContentControl => ({
			id: '4243',
			level: 'inline',
			kind: 'richText',
			sourceXml,
		});
		paragraph.runs.push({
			text: 'Red',
			contentControl: control(`<w:p xmlns:w="${WORD_NS}">${field}</w:p>`),
		});
		const built = '<w:sdt><w:sdtPr><w:id w:val="4243"/></w:sdtPr><w:sdtContent>';
		let xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('INCLUDEPICTURE');
		expect(xml).toContain(built);
		// A w:sdtPr is taken whole only when it holds properties Word reads and nothing else: each
		// element where its schema has it, at any depth, Word's extensions only where Word puts them,
		// no text, no relationship reference, and no id or lock Word does not read. Text that is not
		// XML is not taken either, rather than failing the save.
		const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
		const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
		const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
		const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
		const sdtPr = (inner: string) =>
			`<w:sdtPr xmlns:w="${WORD_NS}" xmlns:w15="${W15}" xmlns:r="${REL}" xmlns:x="urn:x" xmlns:mc="${MC}" xmlns:n="http://schemas.microsoft.com/office/word/not-a-word-namespace">${inner}</w:sdtPr>`;
		for (const sourceXml of [
			sdtPr(`<w:id w:val="4243"/><w:p>${field}</w:p>`),
			sdtPr(`<w:rPr><w:b/>${field}</w:rPr><w:id w:val="4243"/>`),
			sdtPr(`<w:id w:val="4243"/><w:unknownPr/>`),
			sdtPr(`<w:rPr><w:bogus/></w:rPr><w:id w:val="4243"/>`),
			sdtPr(`<w:rPr><w:altChunk r:id="rId9"/></w:rPr><w:id w:val="4243"/>`),
			sdtPr(`<w:id w:val="4243"/><x:anything><x:deep/></x:anything>`),
			sdtPr(`<w:alias w:val="A"><w15:appearance w15:val="hidden"/></w:alias><w:id w:val="4243"/>`),
			sdtPr(`<w:id w:val="4243"/><n:thing/>`),
			sdtPr(`<w:alias w:val="A" x:evil="1"/><w:id w:val="4243"/>`),
			sdtPr(`<w:alias w:val="A" mc:MustUnderstand="x"/><w:id w:val="4243"/>`),
			sdtPr(`<w:id w:val="99999999999"/>`),
			sdtPr(`<w:id w:val="4243"/><w:id w:val="4244"/>`),
			sdtPr(`<w:id w:val="4243"/><w:text/><w:picture/>`),
			sdtPr(`<w:id w:val="4243"/><w:lock w:val="everything"/>`),
			sdtPr(`<w:alias w:val="A"/>text<w:id w:val="4243"/>`),
			'<w:sdtPr>',
		]) {
			paragraph.runs[1]!.contentControl = control(sourceXml);
			xml = await documentXml(await loaded.save(loaded.model));
			expect(xml, sourceXml).toContain(built);
			expect(xml, sourceXml).not.toContain('INCLUDEPICTURE');
		}
		// Properties that are what they say are kept as they were, those the model does not hold
		// and Word's extensions included.
		paragraph.runs[1]!.contentControl = {
			...control(
				`<w:sdtPr xmlns:w="${WORD_NS}" xmlns:w14="${W14}" xmlns:w15="${W15}">` +
					'<w:rPr><w:rStyle w:val="Strong"/><w14:ligatures w14:val="standard"/></w:rPr>' +
					'<w:alias w:val="Kept"/><w:id w:val="4243"/><w:lock w:val="unlocked"/>' +
					'<w:placeholder><w:docPart w:val="DefaultPlaceholder_1"/></w:placeholder>' +
					'<w15:appearance w15:val="hidden"/></w:sdtPr>',
			),
			alias: 'Kept',
		};
		expect(await documentXml(await loaded.save(loaded.model))).toMatch(
			/<w:sdt><w:sdtPr[^>]*><w:rPr><w:rStyle w:val="Strong"\/><w14:ligatures w14:val="standard"\/><\/w:rPr><w:alias w:val="Kept"\/><w:id w:val="4243"\/><w:lock w:val="unlocked"\/><w:placeholder><w:docPart w:val="DefaultPlaceholder_1"\/><\/w:placeholder><w15:appearance w15:val="hidden"\/><\/w:sdtPr><w:sdtContent>/,
		);
	});

	it('builds the properties of a new control only from values Word reads', async () => {
		// A pasted control's kind, id and lock come from another page too: an element named after
		// an unknown kind, an id that is not a 32-bit integer or a lock Word does not know would make
		// the document invalid, so they are left out (a control of an unknown kind is rich text).
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs.push({
			text: 'Red',
			contentControl: {
				id: '99999999999',
				level: 'inline',
				kind: 'bogus' as ContentControl['kind'],
				lock: 'everything' as NonNullable<ContentControl['lock']>,
			},
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:sdt><w:sdtPr/><w:sdtContent>');
		paragraph.runs[1]!.contentControl = {
			id: '-12',
			level: 'inline',
			kind: 'unknown',
			lock: 'contentLocked',
		};
		expect(await documentXml(await loaded.save(loaded.model))).toContain(
			'<w:sdt><w:sdtPr><w:id w:val="-12"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>',
		);
	});

	it('writes only the text values of a new control, a list entry when its text is text', async () => {
		// A model from elsewhere can hold anything: an object where text is written (one whose
		// toString is null cannot even become text) is left out instead of failing the save.
		const crafted = { toString: null } as unknown as string;
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs.push(
			{
				text: 'Red',
				contentControl: {
					id: '41',
					level: 'inline',
					kind: 'dropDownList',
					alias: crafted,
					tag: crafted,
					items: [
						{ display: crafted, value: 'x' },
						{ display: 7, value: 7 } as unknown as { display: string; value: string },
						{ display: 'Fine', value: 'fine' },
					],
				},
			},
			{
				text: 'Today',
				contentControl: {
					id: '42',
					level: 'inline',
					kind: 'date',
					dateFormat: crafted,
					dateLocale: crafted,
					fullDate: crafted,
				},
			},
		);
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(
			'<w:dropDownList><w:listItem w:displayText="x" w:value="x"/><w:listItem w:displayText="7" w:value="7"/><w:listItem w:displayText="Fine" w:value="fine"/></w:dropDownList>',
		);
		expect(xml).not.toContain('w:alias');
		expect(xml).not.toContain('w:tag');
		expect(xml).toContain('<w:date/>');
	});

	it('keeps a full date written in another form when a date control changes', async () => {
		// Another program can write a date without its time; only a value that is not text is left out.
		const loaded = await hand(
			`<w:p><w:sdt><w:sdtPr><w:id w:val="51"/><w:date w:fullDate="2024-01-15"><w:dateFormat w:val="M/d/yyyy"/></w:date></w:sdtPr>` +
				`<w:sdtContent><w:r><w:t>1/15/2024</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
		);
		const control = controlOf(expectParagraph(loaded.model.blocks[0]));
		expect(control.fullDate).toBe('2024-01-15');
		control.dateFormat = 'd MMMM yyyy';
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:date w:fullDate="2024-01-15"><w:dateFormat w:val="d MMMM yyyy"/>');
	});

	it('keeps verbatim properties that hold the attributes later versions of Word write', async () => {
		// Word 2021 and later put a checksum on a data binding, in w:dataBinding and in the
		// w15:dataBinding of a rich text control (w16sdtdh, [MS-DOCX] 2.2.12 and 2.11.2.1), and a UTC
		// date on a formatting revision (w16du, [MS-DOCX] 2.12.2.1): a copy of such a control keeps
		// its properties as Word wrote them, its binding included.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
		const DH = 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash';
		const DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
		const sdtPr = (inner: string) =>
			`<w:sdtPr xmlns:w="${WORD_NS}" xmlns:w15="${W15}" xmlns:w16sdtdh="${DH}" xmlns:w16du="${DU}">${inner}</w:sdtPr>`;
		const revision =
			'<w:rPrChange w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z" w16du:dateUtc="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>';
		paragraph.runs.push({
			text: 'Red',
			contentControl: {
				id: '4245',
				level: 'inline',
				kind: 'text',
				sourceXml: sdtPr(
					`<w:rPr><w:b/>${revision}</w:rPr><w:id w:val="4245"/>` +
						`<w:dataBinding w:xpath="/a" w:storeItemID="{1}" w16sdtdh:storeItemChecksum="Ab12"/><w:text/>`,
				),
			},
		});
		let xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('w16sdtdh:storeItemChecksum="Ab12"');
		expect(xml).toContain('w:xpath="/a"');
		expect(xml).toContain('w16du:dateUtc="2026-01-01T00:00:00Z"');
		paragraph.runs[1]!.contentControl = {
			id: '4245',
			level: 'inline',
			kind: 'richText',
			sourceXml: sdtPr(
				`<w:id w:val="4245"/><w15:dataBinding w:prefixMappings="" w:xpath="/root[1]/data[1]" ` +
					`w:storeItemID="{5D9DE62B-DE87-4789-A4F4-7D834A2F3A10}" w16sdtdh:storeItemChecksum="nqI16g=="/>`,
			),
		};
		xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('w16sdtdh:storeItemChecksum="nqI16g=="');
		expect(xml).toContain('w:xpath="/root[1]/data[1]"');
	});

	it('writes no verbatim properties with a second type element, one of an extension included', async () => {
		// Word's extensions add four types to a w:sdtPr ([MS-DOCX] 2.2.3): a check box and an entity
		// picker in w14, a repeating section and its items in w15. Properties with one of them beside
		// another type, or with an element of an extension twice, are built from the model.
		const loaded = await hand(`<w:p><w:r><w:t>Pick: </w:t></w:r></w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
		const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
		const sdtPr = (inner: string) =>
			`<w:sdtPr xmlns:w="${WORD_NS}" xmlns:w14="${W14}" xmlns:w15="${W15}"><w:id w:val="4246"/>${inner}</w:sdtPr>`;
		paragraph.runs.push({
			text: 'Red',
			contentControl: { id: '4246', level: 'inline', kind: 'text' },
		});
		for (const sourceXml of [
			sdtPr('<w:text/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>'),
			sdtPr('<w:text/><w14:entityPicker/>'),
			sdtPr('<w:text/><w15:repeatingSection/>'),
			sdtPr('<w:text/><w15:repeatingSectionItem/>'),
			sdtPr('<w:text/><w15:appearance w15:val="hidden"/><w15:appearance w15:val="tags"/>'),
		]) {
			paragraph.runs[1]!.contentControl = { id: '4246', level: 'inline', kind: 'text', sourceXml };
			const xml = await documentXml(await loaded.save(loaded.model));
			expect(xml, sourceXml).toContain(
				'<w:sdt><w:sdtPr><w:id w:val="4246"/><w:text/></w:sdtPr><w:sdtContent>',
			);
		}
		// A type of an extension on its own is the control's one type, as Word writes it.
		paragraph.runs[1]!.contentControl = {
			id: '4246',
			level: 'inline',
			kind: 'checkbox',
			checked: false,
			sourceXml: sdtPr(
				'<w14:checkbox><w14:checked w14:val="0"/></w14:checkbox><w15:appearance w15:val="hidden"/>',
			),
		};
		expect(await documentXml(await loaded.save(loaded.model))).toMatch(
			/<w:sdtPr[^>]*><w:id w:val="4246"\/><w14:checkbox><w14:checked w14:val="0"\/><\/w14:checkbox><w15:appearance w15:val="hidden"\/><\/w:sdtPr>/,
		);
	});
});
