import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { LoadedDocument, Paragraph, Table } from '../../model';
import { loadDocx } from '../../parse';
import { LEGACY_DOC_NOTICE, loadLegacyDoc } from '../legacy-doc';

// Word-generated fixtures copied from the ole2 repository (see PROVENANCE.md, "docx legacy .doc").
const fixture = async (name: string) =>
	new Uint8Array(await readFile(new URL(`../__fixtures__/legacy-doc/${name}`, import.meta.url)));
const open = async (name: string): Promise<LoadedDocument> =>
	loadLegacyDoc(await fixture(`${name}.doc`));

/** The model of `loaded` after a save to .docx and a read of that package. */
async function reopened(loaded: LoadedDocument): Promise<LoadedDocument> {
	return loadDocx(await (loaded.saveDocx ? loaded.saveDocx() : loaded.save()));
}

const paragraphs = (loaded: LoadedDocument): Paragraph[] =>
	loaded.model.blocks.filter((block): block is Paragraph => block.type === 'paragraph');
const tables = (loaded: LoadedDocument): Table[] =>
	loaded.model.blocks.filter((block): block is Table => block.type === 'table');
const text = (paragraph: Paragraph): string => paragraph.runs.map((run) => run.text).join('');

describe('legacy .doc conversion', () => {
	it('states what the conversion does and does not do', async () => {
		const loaded = await open('formatting-rich');
		expect(loaded.model.warnings[0]).toBe(LEGACY_DOC_NOTICE);
		expect(LEGACY_DOC_NOTICE).toMatch(/\.docx/);
		expect(LEGACY_DOC_NOTICE).toMatch(/paragraph text edits only/);
	});

	it('maps the stylesheet with inheritance, links and built-in ids', async () => {
		for (const loaded of [
			await open('formatting-styles'),
			await reopened(await open('formatting-styles')),
		]) {
			const styles = loaded.model.paragraphStyles!.styles;
			expect(styles.Normal).toMatchObject({ isDefault: true });
			expect(styles.Heading1).toMatchObject({
				name: 'heading 1',
				basedOn: 'Normal',
				formatting: { keepNext: true, outlineLevel: 1, spacingBeforeTwips: 360 },
			});
			expect(styles.FixtureBody).toMatchObject({
				basedOn: 'Normal',
				formatting: { justification: 'center', indentLeftTwips: 480, spacingAfterTwips: 200 },
			});
			expect(styles.FixtureBodyTwo).toMatchObject({ basedOn: 'FixtureBody' });
			expect(loaded.model.characterStyles!.styles.Heading1Char).toMatchObject({
				linkedStyle: 'Heading1',
			});
			expect(loaded.model.characterStyles!.styles.FixtureEmphasis!.formatting).toMatchObject({
				color: '#008000',
				italic: true,
			});
			expect(paragraphs(loaded).map((p) => p.style)).toEqual([
				'Heading1',
				'FixtureBody',
				'FixtureBodyTwo',
				undefined,
			]);
			expect(paragraphs(loaded)[3]!.runs.find((run) => run.style)).toMatchObject({
				text: 'styled',
				style: 'FixtureEmphasis',
			});
		}
	});

	it('keeps run and paragraph formatting that differs from the style', async () => {
		const loaded = await reopened(await open('formatting-rich'));
		const all = paragraphs(loaded).flatMap((p) => p.runs);
		expect(all.find((run) => run.text === 'bold')).toMatchObject({ bold: true });
		expect(all.find((run) => run.text === 'italic')).toMatchObject({ italic: true });
		expect(all.find((run) => run.text === 'courier')).toMatchObject({ fontFamily: 'Courier New' });
		expect(paragraphs(loaded)[0]).toMatchObject({ align: 'center' });
	});

	it('builds numbering from list definitions and overrides', async () => {
		for (const loaded of [
			await open('formatting-lists'),
			await reopened(await open('formatting-lists')),
		]) {
			const catalog = loaded.model.numberingCatalog!;
			const levels = (id: string) => catalog.abstractNums[catalog.nums[id]!.abstractNumId]!.levels;
			expect(levels('1')[0]).toMatchObject({ numFmt: 'bullet', indentLeftTwips: 720 });
			expect(levels('2')[0]).toMatchObject({ numFmt: 'lowerLetter', lvlText: '%1)', start: 3 });
			expect(levels('3')[1]).toMatchObject({ numFmt: 'decimal', lvlText: '%1.%2' });
			expect(levels('3')[2]).toMatchObject({ numFmt: 'lowerRoman', lvlText: '(%3)' });
			expect(paragraphs(loaded).map((p) => p.numbering)).toEqual([
				{ numId: 1, level: 0 },
				{ numId: 1, level: 0 },
				{ numId: 2, level: 0 },
				{ numId: 2, level: 0 },
				{ numId: 3, level: 0 },
				{ numId: 3, level: 1 },
				{ numId: 3, level: 2 },
				{ numId: 3, level: 0 },
				undefined,
			]);
		}
	});

	it('maps merged cells, row properties and widths of a table', async () => {
		for (const loaded of [
			await open('structure-table-merged'),
			await reopened(await open('structure-table-merged')),
		]) {
			const table = tables(loaded)[0]!;
			expect(table.grid).toEqual([3120, 3120, 3120]);
			expect(table.rows.map((row) => row.length)).toEqual([2, 3, 3]);
			expect(table.rows[0]![0]).toMatchObject({ gridSpan: 2, widthTwips: 6240 });
			expect(table.rows[0]![0]!.paragraphs.map(text)).toEqual(['R1C1', 'R1C2']);
			expect(table.rows[1]![2]).toMatchObject({ verticalMerge: 'restart' });
			expect(table.rows[2]![2]).toMatchObject({ verticalMerge: 'continue' });
			expect(table.rowProperties?.[1]).toMatchObject({ cantSplit: true });
		}
	});

	it('keeps borders, shading, row height, header rows and alignment of a simple table', async () => {
		const zip = await JSZip.loadAsync(await (await open('structure-table-simple')).saveDocx!());
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:tblBorders>');
		expect(xml).toContain('w:fill="FF0000"');
		expect(xml).toContain('<w:tblHeader/>');
		expect(xml).toMatch(/<w:trHeight w:val="600" w:hRule="exact"\/>/);
		expect(xml).toContain('<w:vAlign w:val="center"/>');
		expect(xml).toMatch(/<w:tblPr>.*<w:jc w:val="center"\/>/);
	});

	it('maps a nested table inside a cell', async () => {
		const loaded = await reopened(await open('structure-table-nested'));
		const outer = tables(loaded)[0]!;
		expect(outer.rows[0]![1]!.nestedTables?.[0]?.rows.flat().map((cell) => cell.text)).toEqual([
			'In11',
			'In12',
			'In21',
			'In22',
		]);
	});

	it('maps section geometry, break types and columns', async () => {
		for (const loaded of [
			await open('structure-sections'),
			await reopened(await open('structure-sections')),
		]) {
			const sections = loaded.model.sections!;
			expect(sections.map((s) => s.type)).toEqual([
				'nextPage',
				'oddPage',
				'nextPage',
				'continuous',
			]);
			expect(sections[1]).toMatchObject({
				orientation: 'landscape',
				pageWidthTwips: 15840,
				pageHeightTwips: 12240,
				marginLeftTwips: 1080,
				marginRightTwips: 1800,
				marginTopTwips: 1200,
				marginBottomTwips: 1320,
				gutterTwips: 360,
				headerDistanceTwips: 600,
				footerDistanceTwips: 800,
				columns: { count: 2, spacingTwips: 600 },
			});
		}
	});

	it('maps header and footer variants per section and inherits them', async () => {
		for (const loaded of [
			await open('structure-headers'),
			await reopened(await open('structure-headers')),
		]) {
			const [first, second] = loaded.model.sections!;
			expect(first!.titlePage).toBe(true);
			const textOf = (part: { blocks: Paragraph[] | unknown[] } | undefined) =>
				(part?.blocks as Paragraph[] | undefined)?.map(text);
			expect(textOf(first!.headers!.default)).toEqual(['Odd header 1']);
			expect(textOf(first!.headers!.first)).toEqual(['First header 1']);
			expect(textOf(first!.headers!.even)).toEqual(['Even header 1']);
			expect(textOf(first!.footers!.default)).toEqual(['Odd footer 1']);
			expect(textOf(first!.footers!.first)).toEqual(['First footer 1']);
			expect(textOf(first!.footers!.even)).toEqual(['Even footer 1']);
			expect(textOf(second!.headers!.default)).toHaveLength(1);
			// The second section writes no footer, so it inherits the first section's.
			expect(second!.footers?.default).toBeUndefined();
		}
	});

	it('keeps complex fields with their instruction and result', async () => {
		const loaded = await reopened(await open('header-field'));
		const field = paragraphs(loaded)[1]!.runs.find((run) => run.field);
		expect(field).toMatchObject({
			text: 'Fixture field',
			field: { instr: 'QUOTE "Fixture field"' },
		});
		const header = loaded.model.sections![0]!.headers!.default!.blocks[0] as Paragraph;
		expect(header.runs.find((run) => run.field)?.field?.instr).toBe('PAGE  \\* MERGEFORMAT');
	});

	it('maps footnotes, endnotes and comments with author and date', async () => {
		for (const loaded of [
			await open('structure-notes'),
			await reopened(await open('structure-notes')),
		]) {
			expect(
				loaded.model.footnotes!.map((n) => (n.blocks as Paragraph[]).map(text).join('').trim()),
			).toEqual(['First footnote text', 'Second footnote text']);
			expect(loaded.model.endnotes!).toHaveLength(2);
			expect(loaded.model.comments!.map((c) => [c.author, c.initials, c.text])).toEqual([
				['Alice Author', 'AA', 'First comment text'],
				['Bob Builder', 'BB', 'Second comment text'],
			]);
			expect(loaded.model.comments![0]!.date).toMatch(/^\d{4}-\d{2}-\d{2}T/);
			const references = paragraphs(loaded)[0]!.runs.filter((run) => run.noteReference);
			expect(references.map((run) => run.noteReference!.kind)).toEqual([
				'footnote',
				'footnote',
				'endnote',
				'endnote',
			]);
		}
	});

	it('maps hyperlinks, bookmarks and pictures, keeping the image bytes', async () => {
		const first = await open('objects');
		const loaded = await reopened(first);
		const links = paragraphs(loaded).flatMap((p) => p.runs.filter((run) => run.link));
		expect(links.map((run) => run.link)).toEqual([
			{ href: 'https://example.com/path?x=1', tooltip: 'Example tip' },
			{ anchor: 'BM_Target', tooltip: 'Local tip' },
		]);
		const images = paragraphs(loaded).flatMap((p) => p.runs.filter((run) => run.image?.partName));
		const media = [...(loaded.media?.values() ?? [])];
		const png = new Uint8Array(await fixture('objects-input-a.png'));
		const jpeg = new Uint8Array(await fixture('objects-input-b.jpg'));
		expect(media.some((bytes) => Buffer.compare(bytes, png) === 0)).toBe(true);
		expect(media.some((bytes) => Buffer.compare(bytes, jpeg) === 0)).toBe(true);
		const placements = images
			.filter((run) => run.image?.anchored)
			.map((run) => run.image!.placement);
		expect(placements.map((p) => p?.wrap)).toContain('square');
		expect(images.some((run) => run.image?.contentType === 'image/x-emf')).toBe(true);
		expect(first.model.warnings.join('\n')).toMatch(
			/OLE objects are shown as their preview picture/,
		);
	});

	it('maps floating text boxes', async () => {
		const loaded = await reopened(await open('structure-textbox'));
		const boxes = paragraphs(loaded).flatMap((p) => p.runs.filter((run) => run.image?.textBoxText));
		expect(boxes.map((run) => run.image!.textBoxText)).toEqual([
			['Text box one'],
			['Text box two'],
		]);
	});

	it('round-trips every fixture through a second save without loss of text', async () => {
		const names = [
			'formatting-lists',
			'formatting-rich',
			'formatting-styles',
			'header-field',
			'objects',
			'structure-headers',
			'structure-notes',
			'structure-sections',
			'structure-table-merged',
			'structure-table-nested',
			'structure-table-simple',
			'structure-textbox',
		];
		for (const name of names) {
			const first = await open(name);
			const second = await reopened(first);
			const third = await reopened(second);
			expect(JSON.stringify(third.model.blocks.map((b) => b.type))).toBe(
				JSON.stringify(first.model.blocks.map((b) => b.type)),
			);
			expect(paragraphs(third).map(text)).toEqual(paragraphs(first).map(text));
		}
	});

	it('keeps the binary .doc on save: unchanged bytes, and refuses edits it cannot write', async () => {
		const bytes = await fixture('structure-table-simple.doc');
		const loaded = await loadLegacyDoc(bytes);
		expect(await loaded.save()).toEqual(bytes);
		const cell = tables(loaded)[0]!.rows[0]![0]!.paragraphs[0]!;
		cell.runs[0]!.text += ' edited';
		await expect(loaded.save()).rejects.toThrow(/saveDocx/);
		const docx = await loaded.saveDocx!();
		expect(Array.from(docx.slice(0, 2))).toEqual([0x50, 0x4b]);
	});
});
