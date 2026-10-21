import { describe, expect, it } from 'vitest';
import { formatAddress, parseRange, type CellRange } from '../address';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import type { Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { compareWorkbooks } from './compare';

const R = (ref: string): CellRange => {
	const range = parseRange(ref);
	if (!range) throw new Error(ref);
	return range;
};

/** A workbook touching every compared category, built through the edit API. */
function sample(): Workbook {
	const workbook = createWorkbook({ sheets: ['Inputs', 'Report', 'Lookup'] });
	const s = createEditSession(workbook);
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Item', 'Qty', 'Price', 'Total'],
		['Bolt', 4, 0.25, null],
		['Nut', 10, 0.1, null],
		['Washer', 25, 0.05, null],
	]);
	for (let row = 1; row <= 3; row++) s.setCellInput(0, row, 3, `=B${row + 1}*C${row + 1}`);
	s.setCellInput(0, 5, 3, '=SUM(D2:D4)');
	s.setCellInput(0, 6, 0, "'00123");
	s.setCellValue(0, 7, 0, true);
	s.applyStyle(0, [R('A1:D1')], {
		font: { bold: true },
		fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFDDEEFF' } },
	});
	s.applyStyle(0, [R('C2:D6')], { numFmt: '0.00' });
	s.applyStyle(0, [R('A6:D6')], {
		border: { top: { style: 'thin' } },
		alignment: { horizontal: 'right' },
	});
	s.applyStyle(0, [R('B2:B4')], { protection: { locked: false } });
	s.merge(0, R('A10:C10'), 'merge');
	s.setCellValue(0, 9, 0, 'Notes');
	s.setDefinedName({ name: 'Total', formula: 'Inputs!$D$6' });
	s.setDefinedName({ name: 'Local', formula: 'Report!$A$1', localSheet: 1 });
	s.createTable(0, R('A1:D4'), true, 'TableStyleLight9');
	s.setDataValidation(
		0,
		{ ranges: [], type: 'whole', operator: 'between', formula1: '0', formula2: '100' },
		R('B2:B4'),
	);
	s.addConditionalFormat(0, {
		ranges: [R('D2:D4')],
		rules: [
			{
				type: 'cellIs',
				operator: 'greaterThan',
				formulas: ['1'],
				style: { font: { italic: true } },
				priority: 1,
			},
		],
	});
	s.setCellInput(1, 0, 0, '=Total*2');
	s.setCellInput(1, 1, 0, '=SEQUENCE(3)');
	s.setSheetState(2, 'hidden');
	return workbook;
}

const reload = async (workbook: Workbook) => loadXlsx(await saveXlsx(workbook));

describe('compareWorkbooks round trip', () => {
	it('finds no differences between a workbook and its saved and reloaded copy', async () => {
		const original = sample();
		const reloaded = await reload(original);
		const result = compareWorkbooks(original, reloaded);
		expect(result.count).toBe(0);
		expect(result.sheets.map((sheet) => [sheet.name, sheet.matchedBy])).toEqual([
			['Inputs', 'name'],
			['Report', 'name'],
			['Lookup', 'name'],
		]);
		expect(compareWorkbooks(reloaded, await reload(reloaded)).count).toBe(0);
	});

	it('ignores formats renumbered by a save and reload', async () => {
		const original = sample();
		// A second copy of the default format at index 1, spelled with an explicit `bold: false`:
		// the reader merges equal formats, so every later id moves down by one after the reload.
		const base = original.styles[0]!;
		original.styles.splice(1, 0, { ...base, font: { ...base.font, bold: false } });
		for (const sheet of original.sheets)
			for (const cells of sheet.rows.values())
				for (const cell of cells.values()) if (cell.styleId) cell.styleId += 1;
		const reloaded = await reload(original);
		const header = (book: Workbook) => getCell(book.sheets[0]!, 0, 0)?.styleId;
		expect(header(reloaded)).toBe((header(original) ?? 0) - 1);
		expect(compareWorkbooks(original, reloaded).count).toBe(0);
	});

	it('reports exactly the edits made after a reload', async () => {
		const original = sample();
		const edited = await reload(original);
		const s = createEditSession(edited);
		s.setCellValue(0, 1, 1, 8);
		s.applyStyle(0, [R('A8')], { font: { color: { rgb: 'FFFF0000' } } });
		s.renameSheet(1, 'Summary');
		const result = compareWorkbooks(original, edited);
		const cells = result.sheets.flatMap((sheet) =>
			sheet.cells.map((d) => `${sheet.name}!${formatAddress(d.address)} ${d.kind}`),
		);
		expect(result.workbook.map((d) => d.kind)).toEqual(['sheetRenamed', 'nameChanged']);
		expect(cells).toEqual([
			'Inputs!B2 enteredValue',
			'Inputs!D2 calculatedValue',
			'Inputs!D6 calculatedValue',
			'Inputs!A8 format',
			'Summary!A1 calculatedValue',
		]);
	});
});
