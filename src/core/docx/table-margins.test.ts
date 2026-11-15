import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx, type Table } from './index';
import { resolveTableCellMargins } from './table-margins';
import { parseTableStyleCatalog } from './table-styles';
import { layoutDocumentModel } from './layout/layout';
import type { LayoutTableBox } from './layout/result';
import type { TextMeasurer } from './layout/measure';
import { signedTwips } from './units';

const fixtures = new URL('./__fixtures__/table-autofit/', import.meta.url);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const style = (id: string, inner: string, extra = '') =>
	`<w:style w:type="table" ${extra} w:styleId="${id}"><w:name w:val="${id}"/>${inner}</w:style>`;
const mar = (left: number, right: number) =>
	`<w:tblPr><w:tblCellMar><w:left w:w="${left}" w:type="dxa"/><w:right w:w="${right}" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
const catalog = (inner: string) =>
	parseTableStyleCatalog(`<w:styles xmlns:w="${W}">${inner}</w:styles>`);

describe('resolveTableCellMargins', () => {
	const styles = catalog(
		style('Plain', mar(50, 50), 'w:default="1"') +
			style('Base', mar(200, 300)) +
			style('Child', '<w:basedOn w:val="Base"/>'),
	);

	it('takes the table, then its style chain, then the default table style, side by side', () => {
		expect(resolveTableCellMargins({}, styles)).toMatchObject({ left: 50, right: 50 });
		expect(resolveTableCellMargins({ style: 'Child' }, styles)).toMatchObject({
			left: 200,
			right: 300,
		});
		expect(
			resolveTableCellMargins({ style: 'Child', cellMargins: { left: signedTwips(20) } }, styles),
		).toMatchObject({ left: 20, right: 300, top: 0, bottom: 0 });
	});

	it('falls back to 10 twips with no table style, and to 108 with no styles part', () => {
		expect(resolveTableCellMargins({}, catalog(''))).toEqual({
			top: 0,
			right: 10,
			bottom: 0,
			left: 10,
		});
		expect(resolveTableCellMargins({}, undefined)).toEqual({
			top: 0,
			right: 108,
			bottom: 0,
			left: 108,
		});
	});

	it('stops at a style inheritance cycle', () => {
		const cyclic = catalog(
			style('A', '<w:basedOn w:val="B"/>') + style('B', `<w:basedOn w:val="A"/>${mar(30, 40)}`),
		);
		expect(resolveTableCellMargins({ style: 'A' }, cyclic)).toMatchObject({ left: 30, right: 40 });
	});
});

/** Courier New, the fixtures' only font: every character advances 1229/2048 em. */
const courier: TextMeasurer = {
	widthOf: (text, font) => [...text].length * font.sizePx * (1229 / 2048),
	lineHeightOf: (font) => font.sizePx * 1.1328,
};

describe('cell margins in documents Word saved', () => {
	it.each(['style-none', 'style-default', 'style-chain', 'style-own'])(
		'%s lays its AutoFit columns out on the grid Word wrote',
		async (name) => {
			const bytes = readFileSync(new URL(`${name}.word.docx`, fixtures));
			const { model } = await loadDocx(new Uint8Array(bytes));
			const table = model.blocks.find((block): block is Table => block.type === 'table')!;
			const grid = table.grid!;
			const result = layoutDocumentModel(model, courier);
			const box = result.pages[0]!.columns[0]!.blocks.find(
				(block): block is LayoutTableBox => block.kind === 'table',
			)!;
			const widths = box.rows[0]!.geometry!.map((cell) => Math.round(cell.widthPx * 15));
			for (const [index, width] of widths.entries())
				expect(Math.abs(width - grid[index]!), `column ${index}`).toBeLessThanOrEqual(1);
		},
	);
});
