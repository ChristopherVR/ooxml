import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { tableCellViews } from '../ui/table-cell-views';
import { createFakeMeasurer } from './measure';
import { layoutDocumentModel } from './layout';
import type { LayoutParagraphBox, LayoutResult, LayoutTableBox } from './result';

// The generated documents of the table reports (#320 to #323); see the folder's README for what
// Word 16 does with each of them.
async function load(name: string) {
	const url = new URL(`./fixtures/table-reports/${name}.docx`, import.meta.url);
	return (await loadDocx(new Uint8Array(await readFile(url)))).model;
}

const tables = (result: LayoutResult) =>
	result.pages.flatMap((page, pageIndex) =>
		page.columns.flatMap((column) =>
			column.blocks
				.filter((block): block is LayoutTableBox => block.kind === 'table')
				.map((table) => ({ page, pageIndex, column, table })),
		),
	);
const cellParagraphs = (table: LayoutTableBox): LayoutParagraphBox[] =>
	table.rows.flatMap((row) => row.cells.flat());
const rowWidth = (table: LayoutTableBox) =>
	Math.max(
		...table.rows.map((row) =>
			(row.geometry ?? []).reduce((right, cell) => Math.max(right, cell.xPx + cell.widthPx), 0),
		),
	);

describe('table report fixtures', () => {
	it.each(['t02-autofit-12-colonnes', 't03-mots-insecables'])(
		'keeps the auto-fit table of %s inside the text column, as Word does (#320)',
		async (name) => {
			const result = layoutDocumentModel(await load(name), createFakeMeasurer());
			const [placed] = tables(result);
			expect(placed).toBeDefined();
			const { column, table } = placed!;
			// Without compatibility settings Word lays the table out in its legacy mode: the first
			// cell's left margin (108 twips) hangs into the page margin so its text lines up.
			const overhang = 108 / 15;
			expect((table.xPx ?? 0) + rowWidth(table)).toBeLessThanOrEqual(column.widthPx + overhang + 1);
		},
	);

	it('breaks the unbreakable tokens of t03 inside their cells (#320)', async () => {
		const result = layoutDocumentModel(await load('t03-mots-insecables'), createFakeMeasurer());
		const { table } = tables(result)[0]!;
		// The URL, IBAN and reference cells each wrap onto more than one line, as in Word's PDF.
		for (const row of table.rows.slice(1))
			expect(row.cells[1]!.reduce((lines, p) => lines + p.lines.length, 0)).toBeGreaterThan(1);
	});

	it('resolves Table Grid borders for the header and footer tables of t10 (#321)', async () => {
		const model = await load('t10-tableau-entete-pied');
		const section = model.sections![0]!;
		for (const story of [section.headers?.default, section.footers?.default]) {
			const table = story!.blocks.find((block) => block.type === 'table');
			expect(table?.type).toBe('table');
			if (table?.type !== 'table') continue;
			for (const view of tableCellViews(table, model.tableStyles).flat()) {
				const borders = JSON.parse(view.attrs.borders!) as Record<string, { style?: string }>;
				for (const side of ['top', 'bottom', 'left', 'right'])
					expect(borders[side]?.style).toBe('single');
			}
		}
	});

	it('draws a repeating header row once when the table starts at the top of a page (#324)', async () => {
		const result = layoutDocumentModel(
			await load('t24-entete-repete-haut-de-page'),
			createFakeMeasurer(),
		);
		// Word 16: two pages; the table, header row included, is drawn once on page 2.
		expect(result.pages).toHaveLength(2);
		const placed = tables(result);
		expect(placed.map((entry) => entry.pageIndex)).toEqual([1]);
		const { table } = placed[0]!;
		expect(table.rows.map((row) => row.repeated)).toEqual(Array(6).fill(false));
		expect(table.yPx).toBe(0);
	});

	it('splits a keep-together row taller than a page across pages by default, as LibreOffice does (#323)', async () => {
		const result = layoutDocumentModel(
			await load('t14-ligne-plus-haute-que-page'),
			createFakeMeasurer(),
		);
		// Default ('split', LibreOffice behaviour): ignore cantSplit for rows taller than a page so
		// no paragraphs are clipped or lost.
		const placed = tables(result);
		expect(placed.length).toBeGreaterThan(1);
		expect(placed.flatMap((entry) => cellParagraphs(entry.table))).toHaveLength(80);
		for (const { page, table } of placed)
			expect(table.yPx + table.heightPx).toBeLessThanOrEqual(
				page.heightPx - page.marginTopPx - page.marginBottomPx + 1,
			);
		expect(result.approximations).not.toContain(
			'A table row marked "keep row together" is taller than an empty page/column.',
		);
		const lastPage = result.pages.at(-1)!.columns[0]!.blocks;
		expect(lastPage.at(-1)!.kind).toBe('paragraph');
	});

	it('can clip a keep-together row taller than a page when configured for Word behaviour (#323)', async () => {
		const result = layoutDocumentModel(
			await load('t14-ligne-plus-haute-que-page'),
			createFakeMeasurer(),
			{ oversizedCantSplitRow: 'clip' },
		);
		// Word 16: the row moves to page 2 whole and is cut at the bottom of that page; page 3
		// holds only the paragraph after the table.
		expect(result.pages).toHaveLength(3);
		const placed = tables(result);
		expect(placed.map((entry) => entry.pageIndex)).toEqual([1]);
		expect(cellParagraphs(placed[0]!.table)).toHaveLength(80);
		const last = result.pages[2]!.columns[0]!.blocks;
		expect(last).toHaveLength(1);
		expect(last[0]!.kind).toBe('paragraph');
		expect(result.approximations).toContain(
			'A table row marked "keep row together" is taller than an empty page/column.',
		);
	});
});
