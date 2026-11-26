import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx, twips, type Table } from './index';
import {
	autoFitTable,
	repeatTableRows,
	resizeTableColumns,
	resizeTableRows,
} from './table-layout-edit';
import {
	evaluateTableFormula,
	sortTableRows,
	tableFormulaRuns,
	tableToText,
} from './table-data-edit';
import { TABLE_STYLE_PRESETS, installTableStylePreset } from './table-style-presets';
import { expectTable } from './test-support/access';

function makeTable(): Table {
	return {
		type: 'table',
		id: 'table',
		grid: [twips(1800), twips(2700), twips(3900)],
		rows: [
			['12', '6', '50'],
			['2', 'result', '3'],
			['8', '4', '70'],
		].map((values, ri) =>
			values.map((text, ci) => ({
				widthPercent: 1 / values.length,
				paragraphs: [
					{ type: 'paragraph', id: `cell-${ri}-${ci}`, runs: [{ text, bold: ci === 1 }] },
				],
			})),
		),
	};
}

describe('Word table contextual operations', () => {
	it('round trips sizing and each AutoFit mode through new and existing package writes', async () => {
		const model = createDocument();
		const resized = resizeTableColumns(makeTable(), 0, 1, 8400);
		model.blocks = [resizeTableRows(resized, 0, 1, [250, 370, 200])];
		let document = await loadDocx(await saveDocx(model));
		let table = expectTable(document.model.blocks[0]);
		expect(table.layout).toBe('fixed');
		expect(table.grid).toEqual([2250, 2250, 3900]);
		expect(table.rows[1]?.map((cell) => cell.widthTwips)).toEqual(table.grid);
		expect(table.rowProperties?.map((row) => row.heightTwips)).toEqual([310, 310, undefined]);
		for (const mode of ['window', 'contents', 'fixed'] as const) {
			document.model.blocks[0] = autoFitTable(table, mode);
			document = await loadDocx(await document.save());
			table = expectTable(document.model.blocks[0]);
			expect(table.layout).toBe(mode === 'fixed' ? 'fixed' : 'autofit');
			expect(table.widthPercent).toBe(mode === 'window' ? 1 : undefined);
			expect(
				table.rows[0]?.every(
					(cell) => cell.widthTwips === undefined && cell.widthPercent === undefined,
				),
			).toBe(true);
		}
	});

	it('persists the table style, all look flags, contiguous headers and cell alignment', async () => {
		const model = createDocument();
		const table = repeatTableRows(makeTable(), 0, 1);
		const preset = TABLE_STYLE_PRESETS.find(({ id }) => id === 'GridTable4-Accent1')!;
		table.style = preset.id;
		table.look = {
			firstRow: true,
			lastRow: false,
			firstColumn: false,
			lastColumn: true,
			noHBand: false,
			noVBand: true,
		};
		table.rows[0]![0]!.verticalAlign = 'bottom';
		model.blocks = [table];
		model.tableStyles = installTableStylePreset(undefined, preset);
		const reopened = await loadDocx(await saveDocx(model));
		const loaded = expectTable(reopened.model.blocks[0]);
		expect(loaded.style).toBe(preset.id);
		expect(loaded.look).toEqual(table.look);
		expect(loaded.rowProperties?.map((row) => row.header)).toEqual([true, true, undefined]);
		expect(loaded.rows[0]?.[0]?.verticalAlign).toBe('bottom');
		expect(reopened.model.tableStyles?.styles[preset.id]?.conditional.firstRow?.shadingFill).toBe(
			'#4472C4',
		);
		expect(repeatTableRows(makeTable(), 1, 1).rowProperties).toBeUndefined();
		expect(repeatTableRows(loaded, 0, 0).rowProperties?.map((row) => row.header)).toEqual([
			false,
			false,
			undefined,
		]);
	});

	it.each([
		['=SUM(ABOVE)', '6'],
		['=SUM(BELOW)', '4'],
		['=PRODUCT(LEFT,RIGHT)', '6'],
		['=AVERAGE(ABOVE, BELOW)', '5'],
		['=COUNT(LEFT;RIGHT)', '2'],
		['=MAX(LEFT,BELOW)', '4'],
		['=MIN(RIGHT,ABOVE)', '3'],
	])('evaluates and round trips positional formula %s', async (instruction, value) => {
		const model = createDocument();
		const table = makeTable();
		expect(evaluateTableFormula(table, 1, 1, instruction)).toBe(value);
		const runs = tableFormulaRuns(table, 1, 1, instruction)!;
		table.rows[1]![1]!.paragraphs[0]!.runs = runs;
		model.blocks = [table];
		const loaded = expectTable((await loadDocx(await saveDocx(model))).model.blocks[0]);
		expect(loaded.rows[1]?.[1]?.paragraphs[0]?.runs).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ fieldCode: ` ${instruction} ` }),
				expect.objectContaining({
					text: value,
					field: expect.objectContaining({ instr: instruction }),
				}),
			]),
		);
	});

	it('sorts whole rows stably and exports converted text with the inline formatting', async () => {
		const table = makeTable();
		table.rowProperties = [{ header: true }, { cantSplit: true }, { heightTwips: twips(400) }];
		expect(evaluateTableFormula(table, 1, 1, '=SUM(ABOVE)')).toBe('0');
		expect(evaluateTableFormula(table, 1, 1, '=IF(1,2,3)')).toBeUndefined();
		const sorted = sortTableRows(table, {
			column: 0,
			order: 'descending',
			type: 'number',
			headerRows: 1,
			locale: 'en',
		});
		expect(sorted.rows).toEqual([table.rows[0], table.rows[2], table.rows[1]]);
		expect(sorted.rowProperties).toEqual([
			table.rowProperties[0],
			table.rowProperties[2],
			table.rowProperties[1],
		]);
		const model = createDocument();
		model.blocks = tableToText(sorted, 'tabs');
		const { model: result } = await loadDocx(await saveDocx(model));
		const first = result.blocks[0];
		expect(first?.type).toBe('paragraph');
		if (first?.type !== 'paragraph') return;
		expect(first.runs.map((run) => run.text).join('')).toBe('12\t6\t50');
		expect(first.runs.find((run) => run.text === '6')?.bold).toBe(true);
	});
});
