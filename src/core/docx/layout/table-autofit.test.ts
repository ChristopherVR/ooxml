import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx, type Table } from '../index';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';
import type { LayoutTableBox } from './result';
import { resolveTableAutofit } from './table-autofit';
import type { LayoutTable } from './input';

const fixtures = new URL('../__fixtures__/table-autofit/', import.meta.url);
const recorded = JSON.parse(readFileSync(new URL('word-grids.json', fixtures), 'utf8')) as {
	cases: Record<string, { mode: number; grid: number[] }>;
};

/** Courier New, the fixtures' only font: every character advances 1229/2048 em. */
const courier: TextMeasurer = {
	widthOf: (text, font) =>
		[...text].filter((ch) => ch !== '\t' && ch !== '\n').length * font.sizePx * (1229 / 2048),
	lineHeightOf: (font) => font.sizePx * 1.1328,
};

async function laidOutTable(name: string) {
	const { model } = await loadDocx(new Uint8Array(readFileSync(new URL(`${name}.docx`, fixtures))));
	const table = model.blocks.find((block): block is Table => block.type === 'table')!;
	const result = layoutDocumentModel(model, courier);
	const box = result.pages[0]!.columns[0]!.blocks.find(
		(block): block is LayoutTableBox => block.kind === 'table',
	)!;
	return { table, box };
}

describe('AutoFit table columns against the grids Word saved', () => {
	it.each(Object.entries(recorded.cases))('%s', async (name, { grid }) => {
		const { table, box } = await laidOutTable(name);
		const edges = grid.map((_, index) =>
			grid.slice(0, index).reduce((sum, width) => sum + width, 0),
		);
		for (const [rowIndex, row] of box.rows.entries()) {
			let start = 0;
			for (const [cellIndex, geometry] of (row.geometry ?? []).entries()) {
				const span = table.rows[rowIndex]![cellIndex]!.gridSpan ?? 1;
				const expectedX = edges[start]!;
				const expectedWidth = grid.slice(start, start + span).reduce((sum, w) => sum + w, 0);
				expect(
					Math.abs(geometry.xPx * 15 - expectedX),
					`row ${rowIndex} cell ${cellIndex} x`,
				).toBeLessThanOrEqual(1.01);
				expect(
					Math.abs(geometry.widthPx * 15 - expectedWidth),
					`row ${rowIndex} cell ${cellIndex} width`,
				).toBeLessThanOrEqual(1.01);
				start += span;
			}
		}
	});
});

describe('resolveTableAutofit', () => {
	const paragraph = (id: string, text: string) => ({
		kind: 'paragraph' as const,
		id,
		runs: [{ text, fontSizePt: 11 }],
	});
	const padding = { top: 0, right: 0, bottom: 0, left: 0 };
	const table = (texts: string[], autofit = true): LayoutTable => ({
		kind: 'table',
		id: 't',
		widthPx: 300,
		rows: [
			{
				cells: texts.map((text, index) => ({
					paragraphs: [paragraph(`p${index}`, text)],
					widthPx: 100,
					xPx: index * 100,
					padding,
				})),
			},
		],
		...(autofit ? { autofit: { columns: texts.length } } : {}),
	});

	it('leaves fixed-layout tables untouched', () => {
		const fixed = table(['a', 'bb'], false);
		expect(resolveTableAutofit(fixed, 600, courier)).toBe(fixed);
	});

	it('sizes columns to their content and keeps one object per text column width', () => {
		const source = table(['A', 'Hello world']);
		const resolved = resolveTableAutofit(source, 600, courier);
		const char = ((11 * 96) / 72) * (1229 / 2048);
		const cells = resolved.rows[0]!.cells;
		expect(Math.abs(cells[0]!.widthPx! - char)).toBeLessThan(0.07);
		expect(Math.abs(cells[1]!.widthPx! - 11 * char)).toBeLessThan(0.07);
		expect(cells[1]!.xPx).toBe(cells[0]!.widthPx);
		expect(resolveTableAutofit(source, 600, courier)).toBe(resolved);
		expect(resolveTableAutofit(source, 500, courier)).not.toBe(resolved);
		// Paragraph objects are shared, so footnote and keep lookups by paragraph still match.
		expect(cells[0]!.paragraphs[0]).toBe(source.rows[0]!.cells[0]!.paragraphs[0]);
	});

	it('sizes a cell holding a nested AutoFit table from the nested columns', () => {
		const inner = table(['Hello', 'world']);
		const outer: LayoutTable = {
			...table(['', 'x']),
			rows: [
				{
					cells: [
						{ paragraphs: [], blocks: [inner], padding },
						{ paragraphs: [paragraph('x', 'x')], padding },
					],
				},
			],
		};
		const char = ((11 * 96) / 72) * (1229 / 2048);
		const resolved = resolveTableAutofit(outer, 600, courier);
		expect(Math.abs(resolved.rows[0]!.cells[0]!.widthPx! - 10 * char)).toBeLessThan(0.14);
	});

	it('keeps the table inside the text column when an unbreakable word is wider (#320)', () => {
		const char = ((11 * 96) / 72) * (1229 / 2048);
		const resolved = resolveTableAutofit(
			table(['https://example.com/a/very/long/unbroken/address', 'xy']),
			20 * char,
			courier,
		);
		const cells = resolved.rows[0]!.cells;
		expect(resolved.widthPx!).toBeLessThanOrEqual(20 * char + 0.07);
		// The narrow column keeps its minimum; the long token gives way and breaks in its cell.
		expect(Math.abs(cells[1]!.widthPx! - 2 * char)).toBeLessThan(0.07);
		expect(Math.abs(cells[0]!.widthPx! - 18 * char)).toBeLessThan(0.14);
	});

	it('shares the text column between many columns whose minimums do not fit (#320)', () => {
		const words = Array.from({ length: 12 }, (_, index) => `Colonne${index + 1}`);
		const resolved = resolveTableAutofit(table(words), 600, courier);
		const widths = resolved.rows[0]!.cells.map((cell) => cell.widthPx!);
		expect(resolved.widthPx!).toBeLessThanOrEqual(600.07);
		expect(Math.max(...widths) - Math.min(...widths)).toBeLessThan(0.14);
	});

	it('lets a preferred table width below the minimum give way to the minimum', () => {
		const source = {
			...table(['Supercalifragilistic', 'xy']),
			autofit: { columns: 2, widthPx: 20 },
		};
		const char = ((11 * 96) / 72) * (1229 / 2048);
		const resolved = resolveTableAutofit(source, 600, courier);
		expect(resolved.rows[0]!.cells[0]!.widthPx!).toBeGreaterThanOrEqual(20 * char - 0.07);
	});
});
