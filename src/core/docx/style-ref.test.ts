import { describe, expect, it } from 'vitest';
import { computeListLabels, loadDocx } from './index';
import { documentParagraphs, parseStyleReference, styleRefSource, styleTarget } from './style-ref';
import { refreshStyleRefFields } from './style-ref-fields';
import { headingNumberingFixture } from './test-support/heading-numbering-fixture';
import { complexField, paragraph, styleRefFixture } from './test-support/style-ref-fixture';

const load = async (parts: Parameters<typeof styleRefFixture>[0]) =>
	(await loadDocx(await styleRefFixture(parts))).model;

describe('STYLEREF instructions', () => {
	it('reads a quoted style name, its switches and the formats that keep the text', () => {
		expect(parseStyleReference(' STYLEREF "Doc Number" \\* MERGEFORMAT ')).toEqual({
			style: 'Doc Number',
			last: false,
			number: false,
			numberOnly: false,
		});
		expect(parseStyleReference('styleref Title \\l \\* Upper \\* CHARFORMAT')).toMatchObject({
			style: 'Title',
			last: true,
			textCase: 'upper',
			charFormat: true,
		});
		expect(parseStyleReference('STYLEREF 1 \\n \\t')).toMatchObject({
			style: '1',
			number: true,
			numberOnly: true,
		});
		// Word's Edit Field dialog writes the switches before the style.
		expect(parseStyleReference(' STYLEREF \\s "Heading 1" ')).toEqual({
			style: 'Heading 1',
			last: false,
			number: true,
			numberOnly: true,
		});
		expect(parseStyleReference('STYLEREF \\l \\* Upper Title')).toMatchObject({
			style: 'Title',
			last: true,
			textCase: 'upper',
		});
	});

	it('leaves fields it cannot evaluate to the result they were saved with', () => {
		for (const instr of [
			'STYLEREF Title \\p',
			'STYLEREF Title \\r',
			'STYLEREF Title \\w',
			'STYLEREF Title \\# "0.00"',
			'STYLEREF Title \\* Arabic',
			'STYLEREF Doc Number',
			'STYLEREF \\s Doc Number',
			'STYLEREF \\l',
			'STYLEREF',
			'REF Title',
		])
			expect(parseStyleReference(instr), instr).toBeUndefined();
	});
});

describe('STYLEREF styles', () => {
	it('finds a style by name in any case, by heading level, then by id', async () => {
		const model = await load({ body: paragraph('Body') });
		expect(styleTarget(model, 'Doc Number')).toEqual({ kind: 'paragraph', id: 'DocNumber' });
		expect(styleTarget(model, 'doc number')).toEqual({ kind: 'paragraph', id: 'DocNumber' });
		expect(styleTarget(model, 'DocNumber')).toEqual({ kind: 'paragraph', id: 'DocNumber' });
		expect(styleTarget(model, '1')).toEqual({ kind: 'paragraph', id: 'Heading1' });
		expect(styleTarget(model, 'Heading 1')).toEqual({ kind: 'paragraph', id: 'Heading1' });
		expect(styleTarget(model, 'Entry Name')).toEqual({ kind: 'character', id: 'EntryName' });
		expect(styleTarget(model, 'Normal')).toEqual({
			kind: 'paragraph',
			id: 'Normal',
			isDefault: true,
		});
		expect(styleTarget(model, 'Missing')).toBeUndefined();
	});

	it('finds a heading level by the built-in heading name, whatever its id', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		// Word names its built-in headings `heading 1` to `heading 9` in every language and takes
		// the id from the name shown in its interface: `berschrift1` in German, `Titre1` in French.
		const model = await load({
			styles: `<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:default="1" w:styleId="Standard"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Chapter"/></w:style></w:styles>`,
			body: paragraph('Body'),
		});
		expect(styleTarget(model, '1')).toEqual({ kind: 'paragraph', id: 'berschrift1' });
		expect(styleTarget(model, 'Heading 1')).toEqual({ kind: 'paragraph', id: 'berschrift1' });
		// A style named otherwise is not the built-in heading, for Word either.
		expect(styleTarget(model, '2')).toBeUndefined();
		expect(styleTarget(model, 'Chapter')).toEqual({ kind: 'paragraph', id: 'Heading2' });
	});

	it('reads paragraph text without hidden or deleted text, and stretches of a character style', async () => {
		const model = await load({
			body:
				paragraph(
					'',
					'DocNumber',
					'<w:r><w:t>DOC-</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r><w:del w:id="1" w:author="A"><w:r><w:delText>old</w:delText></w:r></w:del><w:r><w:t>7</w:t></w:r>',
				) +
				paragraph(
					'',
					undefined,
					'<w:r><w:rPr><w:rStyle w:val="EntryName"/></w:rPr><w:t>Ash</w:t></w:r><w:r><w:t>, a tree. </w:t></w:r><w:r><w:rPr><w:rStyle w:val="EntryName"/></w:rPr><w:t>Aspen</w:t></w:r>',
				) +
				paragraph('No style'),
		});
		const source = styleRefSource(model);
		const reference = parseStyleReference('STYLEREF "Doc Number"')!;
		expect(source.results(0, reference, styleTarget(model, 'Doc Number')!)).toEqual(['DOC-7']);
		const entries = parseStyleReference('STYLEREF "Entry Name"')!;
		expect(source.results(1, entries, styleTarget(model, 'Entry Name')!)).toEqual(['Ash', 'Aspen']);
		expect(source.results(2, reference, styleTarget(model, 'Doc Number')!)).toEqual([]);
		// A paragraph without a style is in the default paragraph style.
		expect(source.results(2, reference, styleTarget(model, 'Normal')!)).toEqual(['No style']);
	});

	it('leaves out a paragraph a tracked change deletes whole, as it does deleted text', async () => {
		const deleted = (text: string, style: string | undefined, run = '') =>
			`<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}<w:rPr><w:del w:id="9" w:author="A"/></w:rPr></w:pPr><w:del w:id="8" w:author="A"><w:r>${run}<w:delText>${text}</w:delText></w:r></w:del></w:p>`;
		const model = await load({
			body:
				paragraph('Scope', 'Heading1') +
				deleted('Method', 'Heading1') +
				paragraph('', undefined, complexField('STYLEREF 1', ['old'])) +
				// The paragraph mark alone deleted: its text stays, and so does the match.
				`<w:p><w:pPr><w:pStyle w:val="DocNumber"/><w:rPr><w:del w:id="7" w:author="A"/></w:rPr></w:pPr><w:r><w:t>DOC-2</w:t></w:r></w:p>` +
				deleted('Ash', undefined, '<w:rPr><w:rStyle w:val="EntryName"/></w:rPr>'),
			header: paragraph('', undefined, complexField('STYLEREF 1', ['old'])),
		});
		const source = styleRefSource(model);
		const results = (index: number, style: string) =>
			source.results(
				index,
				parseStyleReference(`STYLEREF "${style}"`)!,
				styleTarget(model, style)!,
			);
		expect(results(1, 'heading 1')).toEqual([]);
		expect(results(3, 'Doc Number')).toEqual(['DOC-2']);
		expect(results(4, 'Entry Name')).toEqual([]);
		// The body field reads the heading before the deleted one, and the header takes it on save.
		const refreshed = refreshStyleRefFields(model, { body: true });
		const fieldText = (blocks: readonly unknown[], index: number) =>
			(blocks[index] as { runs: { text: string; field?: unknown }[] }).runs
				.filter((run) => run.field)
				.map((run) => run.text)
				.join('');
		expect(fieldText(refreshed.blocks, 2)).toBe('Scope');
		expect(fieldText(refreshed.sections![0]!.headers!.default!.blocks, 0)).toBe('Scope');
	});

	it('finds hidden text, however it is hidden, and shows nothing for it, as Word does', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const styles = `<w:styles xmlns:w="${w}">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="DocNumber"><w:name w:val="Doc Number"/></w:style>
<w:style w:type="paragraph" w:styleId="Secret"><w:name w:val="Secret"/><w:rPr><w:vanish/></w:rPr></w:style>
<w:style w:type="character" w:styleId="EntryName"><w:name w:val="Entry Name"/></w:style>
<w:style w:type="character" w:styleId="Hidden"><w:name w:val="Hidden"/><w:rPr><w:vanish/></w:rPr></w:style>
<w:style w:type="character" w:styleId="HiddenEntry"><w:name w:val="Hidden Entry"/><w:basedOn w:val="EntryName"/><w:rPr><w:vanish/></w:rPr></w:style>
</w:styles>`;
		const run = (text: string, properties = '') =>
			`<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
		const model = await load({
			styles,
			body:
				// Text hidden through its character style, and through its paragraph style.
				paragraph(
					'',
					'DocNumber',
					run('DOC-') + run('secret', '<w:rStyle w:val="Hidden"/>') + run('7'),
				) +
				paragraph('', 'Secret', run('Hidden title') + run(' shown', '<w:vanish w:val="0"/>')) +
				// A hidden stretch in the character style is found, and shows nothing; hidden text
				// outside it ends a stretch (the way to keep a heading out of the header on its own page).
				paragraph(
					'',
					undefined,
					run('X', '<w:rStyle w:val="EntryName"/><w:vanish/>') +
						run('Y', '<w:vanish/>') +
						run('Ash', '<w:rStyle w:val="EntryName"/>'),
				) +
				paragraph('', undefined, run('Rowan', '<w:rStyle w:val="HiddenEntry"/>')),
		});
		const source = styleRefSource(model);
		const results = (at: number, style: string) =>
			source.results(at, parseStyleReference(`STYLEREF "${style}"`)!, styleTarget(model, style)!);
		expect(results(0, 'Doc Number')).toEqual(['DOC-7']);
		expect(results(1, 'Secret')).toEqual([' shown']);
		expect(results(2, 'Entry Name')).toEqual(['', 'Ash']);
		expect(results(3, 'Hidden Entry')).toEqual(['']);
	});

	it('reads only the numbers and the delimiters between them with \\n \\t', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const model = await load({
			body: paragraph(
				'',
				undefined,
				'<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Scope</w:t></w:r>',
			),
			numbering: `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Part %1"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimalZero"/><w:lvlText w:val="Section %1.%2:"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
		});
		const source = styleRefSource(model);
		const target = styleTarget(model, '1')!;
		expect(source.results(0, parseStyleReference('STYLEREF 1 \\n')!, target)).toEqual([
			'Section 2.01:',
		]);
		expect(source.results(0, parseStyleReference('STYLEREF 1 \\n \\t')!, target)).toEqual(['2.01']);
		// Word writes \s for the chapter number of a caption; it reads as \n \t.
		expect(parseStyleReference('STYLEREF 1 \\s')).toMatchObject({ number: true, numberOnly: true });
		expect(source.results(0, parseStyleReference('STYLEREF 1 \\s')!, target)).toEqual(['2.01']);
	});

	it('keeps the result of a field with \\t when words stand between the levels of its number', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const model = await load({
			body:
				paragraph(
					'',
					undefined,
					'<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Scope</w:t></w:r>',
				) + paragraph('', undefined, complexField('STYLEREF 1 \\n \\t', ['old'])),
			numbering: `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Part %1"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Part %1, Section %2"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
		});
		const source = styleRefSource(model);
		const target = styleTarget(model, '1')!;
		expect(source.results(0, parseStyleReference('STYLEREF 1 \\n')!, target)).toEqual([
			'Part 2, Section 1',
		]);
		// Which of those words Word would drop is not said anywhere, so the saved result stays.
		expect(source.results(0, parseStyleReference('STYLEREF 1 \\n \\t')!, target)).toEqual([
			undefined,
		]);
		const refreshed = refreshStyleRefFields(model, { body: true });
		expect(refreshed.blocks[1]).toBe(model.blocks[1]);
	});

	it('numbers an unnumbered paragraph in the style 0 with \\n and \\s, as Word does', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const heading = (text: string, numId?: string) =>
			paragraph(
				'',
				undefined,
				`<w:pPr><w:pStyle w:val="Heading1"/>${numId === undefined ? '' : `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`}</w:pPr><w:r><w:t>${text}</w:t></w:r>`,
			);
		const model = await load({
			body:
				heading('Scope', '1') +
				// Numbering taken off one heading, and none on another.
				heading('Glossary', '0') +
				heading('Index') +
				paragraph('', undefined, complexField('STYLEREF 1 \\s', ['1'])) +
				heading('Notes', '2'),
			numbering: `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Chapter %1"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="*"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
		});
		const source = styleRefSource(model);
		const target = styleTarget(model, '1')!;
		const number = parseStyleReference('STYLEREF 1 \\n')!;
		const chapter = parseStyleReference('STYLEREF 1 \\s')!;
		expect(source.results(0, number, target)).toEqual(['Chapter 1']);
		expect(source.results(0, chapter, target)).toEqual(['1']);
		// The nearest paragraph in the style is the match, numbered or not: Word shows 0.
		expect(source.results(1, number, target)).toEqual(['0']);
		expect(source.results(2, chapter, target)).toEqual(['0']);
		// What Word shows for a bullet is not documented: the field keeps its saved result.
		expect(source.results(4, number, target)).toEqual([undefined]);
		const refreshed = refreshStyleRefFields(model, { body: true });
		const text = (block: unknown) =>
			(block as { runs: { text: string }[] }).runs.map((run) => run.text).join('');
		expect(text(refreshed.blocks[3])).toBe('0');
	});

	it('numbers a heading in a table nested in a cell like the headings around it', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const heading = (text: string) =>
			paragraph(
				'',
				undefined,
				`<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r>`,
			);
		const cell = (inner: string) =>
			`<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
		const table = (inner: string) =>
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${inner}</w:tr></w:tbl>`;
		const model = await load({
			body:
				heading('Scope') +
				table(cell(paragraph('Outer') + table(cell(heading('Inner'))) + paragraph(''))) +
				paragraph('', undefined, complexField('STYLEREF 1 \\n', ['old'])) +
				heading('Last'),
			numbering: `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Chapter %1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
		});
		// Word numbers a list in reading order, through the tables nested in a cell.
		const labels = computeListLabels(model);
		expect(documentParagraphs(model.blocks).map((p) => labels.get(p.id)?.text ?? null)).toEqual([
			'Chapter 1',
			null,
			'Chapter 2',
			null,
			null,
			'Chapter 3',
		]);
		const refreshed = refreshStyleRefFields(model, { body: true });
		const text = (block: unknown) =>
			(block as { runs: { text: string }[] }).runs.map((run) => run.text).join('');
		expect(text(refreshed.blocks[2])).toBe('Chapter 2');
	});

	it('reads a cell and the tables nested in it in reading order', async () => {
		const cell = (inner: string) =>
			`<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
		const table = (inner: string) =>
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${inner}</w:tr></w:tbl>`;
		const field = paragraph('', undefined, complexField('STYLEREF 1', ['old']));
		const model = await load({
			body: table(
				cell(paragraph('Alpha', 'Heading1') + table(cell(paragraph('Beta', 'Heading1'))) + field),
			),
		});
		const text = (item: { runs: { text: string }[] }) => item.runs.map((run) => run.text).join('');
		expect(documentParagraphs(model.blocks).map(text)).toEqual(['Alpha', 'Beta', 'old']);
		// The nearest heading before the field is the one in the nested table.
		const refreshed = documentParagraphs(refreshStyleRefFields(model, { body: true }).blocks);
		expect(text(refreshed.at(-1)!)).toBe('Beta');
	});

	it('reads the number of a numbered paragraph with \\n, or its text in a case format', async () => {
		const { model } = await loadDocx(await headingNumberingFixture());
		const source = styleRefSource(model);
		const target = styleTarget(model, '2')!;
		expect(source.results(2, parseStyleReference('STYLEREF 2 \\n')!, target)).toEqual(['1.2']);
		expect(source.results(2, parseStyleReference('STYLEREF 2')!, target)).toEqual(['Heading 2']);
		expect(source.results(2, parseStyleReference('STYLEREF 2 \\* Upper')!, target)).toEqual([
			'HEADING 2',
		]);
	});
});
