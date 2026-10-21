import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
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
		});
		expect(parseStyleReference('STYLEREF 1 \\n \\t')).toMatchObject({
			style: '1',
			number: true,
			numberOnly: true,
		});
	});

	it('leaves fields it cannot evaluate to the result they were saved with', () => {
		for (const instr of [
			'STYLEREF Title \\p',
			'STYLEREF Title \\r',
			'STYLEREF Title \\w',
			'STYLEREF 1 \\s',
			'STYLEREF Title \\# "0.00"',
			'STYLEREF Title \\* Arabic',
			'STYLEREF Doc Number',
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
