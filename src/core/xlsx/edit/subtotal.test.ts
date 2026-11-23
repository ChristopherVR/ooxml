import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import type { SubtotalFunction, SubtotalFunctionCode } from './subtotal-fn';

function setupSales() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Region', 'Rep', 'Sales', 'Units'],
		['East', 'Alice', 100, 10],
		['East', 'Bob', 200, 20],
		['West', 'Cara', 50, 5],
		['West', 'Dan', 150, 15],
		['West', 'Eve', 100, 10],
	]);
	return { workbook, session, sheet: workbook.sheets[0]! };
}

describe('applySubtotal and removeSubtotals', () => {
	it('groups rows, inserts SUBTOTAL formulas and Grand Total, sets outline levels, and recalculates', () => {
		const { session, sheet } = setupSales();
		const res = session.applySubtotal(0, parseRange('A1:D6')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [2, 3],
		});

		expect(res.subtotalRows).toEqual([3, 7]);
		expect(res.grandTotalRow).toBe(8);
		expect(res.range).toEqual(parseRange('A1:D9'));

		// Group 1: East Total at row 3
		expect(getCell(sheet, 3, 0)?.value).toBe('East Total');
		expect(getCell(sheet, 3, 2)?.formula).toBe('SUBTOTAL(9,C2:C3)');
		expect(getCell(sheet, 3, 2)?.value).toBe(300);
		expect(getCell(sheet, 3, 3)?.formula).toBe('SUBTOTAL(9,D2:D3)');
		expect(getCell(sheet, 3, 3)?.value).toBe(30);

		// Group 2: West Total at row 7
		expect(getCell(sheet, 7, 0)?.value).toBe('West Total');
		expect(getCell(sheet, 7, 2)?.formula).toBe('SUBTOTAL(9,C5:C7)');
		expect(getCell(sheet, 7, 2)?.value).toBe(300);
		expect(getCell(sheet, 7, 3)?.formula).toBe('SUBTOTAL(9,D5:D7)');
		expect(getCell(sheet, 7, 3)?.value).toBe(30);

		// Grand Total at row 8 (ignores nested SUBTOTAL rows)
		expect(getCell(sheet, 8, 0)?.value).toBe('Grand Total');
		expect(getCell(sheet, 8, 2)?.formula).toBe('SUBTOTAL(9,C2:C8)');
		expect(getCell(sheet, 8, 2)?.value).toBe(600);
		expect(getCell(sheet, 8, 3)?.formula).toBe('SUBTOTAL(9,D2:D8)');
		expect(getCell(sheet, 8, 3)?.value).toBe(60);

		// Row outline levels: header=none, data=2, subtotals=1, grand total=none
		expect(sheet.rowInfo.get(0)?.outlineLevel).toBeUndefined();
		expect(sheet.rowInfo.get(1)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(2)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(3)?.outlineLevel).toBe(1);
		expect(sheet.rowInfo.get(4)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(5)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(6)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(7)?.outlineLevel).toBe(1);
		expect(sheet.rowInfo.get(8)?.outlineLevel).toBeUndefined();
	});

	it.each<[SubtotalFunction, SubtotalFunctionCode, number, number]>([
		['average', 1, 150, 120],
		['count', 2, 2, 5],
		['counta', 3, 2, 5],
		['max', 4, 200, 200],
		['min', 5, 100, 50],
		['product', 6, 20000, 15_000_000_000],
		['stdev', 7, Math.sqrt(5000), Math.sqrt(3250)],
		['stdevp', 8, 50, Math.sqrt(2600)],
		['sum', 9, 300, 600],
		['var', 10, 5000, 3250],
		['varp', 11, 2500, 2600],
	])(
		'supports %s (code %s) by name and by numeric code',
		(name, code, expectedEast, expectedGrand) => {
			for (const fn of [name, code] as const) {
				const { session, sheet } = setupSales();
				const res = session.applySubtotal(0, parseRange('A1:D6')!, {
					atEachChangeIn: 0,
					fn,
					addSubtotalTo: [2],
				});
				expect(getCell(sheet, res.subtotalRows[0]!, 2)?.formula).toBe(`SUBTOTAL(${code},C2:C3)`);
				expect(getCell(sheet, res.subtotalRows[0]!, 2)?.value).toBeCloseTo(expectedEast);
				expect(getCell(sheet, res.grandTotalRow!, 2)?.value).toBeCloseTo(expectedGrand);
			}
		},
	);

	it('supports summaryBelowData: false (Grand Total at top, group subtotals above each group)', () => {
		const { session, sheet } = setupSales();
		const res = session.applySubtotal(0, parseRange('A1:D6')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [2],
			summaryBelowData: false,
		});

		expect(res.grandTotalRow).toBe(1);
		expect(res.subtotalRows).toEqual([2, 5]);
		expect(res.range).toEqual(parseRange('A1:D9'));

		expect(getCell(sheet, 1, 0)?.value).toBe('Grand Total');
		expect(getCell(sheet, 1, 2)?.formula).toBe('SUBTOTAL(9,C3:C9)');
		expect(getCell(sheet, 1, 2)?.value).toBe(600);

		expect(getCell(sheet, 2, 0)?.value).toBe('East Total');
		expect(getCell(sheet, 2, 2)?.formula).toBe('SUBTOTAL(9,C4:C5)');
		expect(getCell(sheet, 2, 2)?.value).toBe(300);

		expect(getCell(sheet, 5, 0)?.value).toBe('West Total');
		expect(getCell(sheet, 5, 2)?.formula).toBe('SUBTOTAL(9,C7:C9)');
		expect(getCell(sheet, 5, 2)?.value).toBe(300);

		expect(sheet.rowInfo.get(1)?.outlineLevel).toBeUndefined();
		expect(sheet.rowInfo.get(2)?.outlineLevel).toBe(1);
		expect(sheet.rowInfo.get(3)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(4)?.outlineLevel).toBe(2);
		expect(sheet.rowInfo.get(5)?.outlineLevel).toBe(1);
		expect(sheet.rowInfo.get(6)?.outlineLevel).toBe(2);
	});

	it('replaces existing subtotals when replaceExisting is true and nests when false', () => {
		const { session, sheet } = setupSales();
		const first = session.applySubtotal(0, parseRange('A1:D6')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [2],
		});
		expect(getCell(sheet, first.grandTotalRow!, 2)?.value).toBe(600);

		// Replace with AVERAGE
		const replaced = session.applySubtotal(0, first.range, {
			atEachChangeIn: 0,
			fn: 'average',
			addSubtotalTo: [2],
			replaceExisting: true,
		});
		expect(replaced.range).toEqual(parseRange('A1:D9'));
		expect(getCell(sheet, replaced.subtotalRows[0]!, 0)?.value).toBe('East Average');
		expect(getCell(sheet, replaced.subtotalRows[0]!, 2)?.value).toBe(150);
		expect(getCell(sheet, replaced.grandTotalRow!, 0)?.value).toBe('Grand Average');
		expect(getCell(sheet, replaced.grandTotalRow!, 2)?.value).toBe(120);

		// Nest MAX without replacing existing AVERAGE subtotals
		const nested = session.applySubtotal(0, replaced.range, {
			atEachChangeIn: 0,
			fn: 'max',
			addSubtotalTo: [2],
			replaceExisting: false,
		});
		// Old subtotals remain alongside new nested subtotals; data rows rise to outline level 3
		expect(nested.range.end.row).toBe(11);
		expect(sheet.rowInfo.get(1)?.outlineLevel).toBe(3);
		expect(getCell(sheet, nested.grandTotalRow!, 0)?.value).toBe('Grand Max');
		expect(getCell(sheet, nested.grandTotalRow!, 2)?.value).toBe(200);
	});

	it('removes all subtotals and clears outline levels with removeSubtotals', () => {
		const { session, sheet } = setupSales();
		const applied = session.applySubtotal(0, parseRange('A1:D6')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [2, 3],
		});

		const removed = session.removeSubtotals(0, applied.range);
		expect(removed.removed).toBe(3);
		expect(removed.range).toEqual(parseRange('A1:D6'));

		// Original rows restored
		expect(getCell(sheet, 1, 0)?.value).toBe('East');
		expect(getCell(sheet, 3, 0)?.value).toBe('West');
		expect(getCell(sheet, 5, 1)?.value).toBe('Eve');
		expect(getCell(sheet, 6, 0)).toBeUndefined();

		// Outline levels cleared
		for (let r = 0; r <= 5; r++) expect(sheet.rowInfo.get(r)?.outlineLevel).toBeUndefined();
	});

	it('supports full undo and redo for applySubtotal and removeSubtotals', () => {
		const { session, sheet } = setupSales();
		const applied = session.applySubtotal(0, parseRange('A1:D6')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [2],
		});
		expect(session.undoLabel()).toBe('Subtotal');
		expect(getCell(sheet, 8, 2)?.value).toBe(600);

		expect(session.undo()).toBe(true);
		expect(getCell(sheet, 5, 1)?.value).toBe('Eve');
		expect(getCell(sheet, 6, 0)).toBeUndefined();
		expect(sheet.rowInfo.get(1)?.outlineLevel).toBeUndefined();

		expect(session.redo()).toBe(true);
		expect(getCell(sheet, 3, 2)?.value).toBe(300);
		expect(getCell(sheet, 8, 2)?.value).toBe(600);
		expect(sheet.rowInfo.get(1)?.outlineLevel).toBe(2);

		session.removeSubtotals(0, applied.range);
		expect(session.undoLabel()).toBe('Remove subtotals');
		expect(getCell(sheet, 6, 0)).toBeUndefined();

		expect(session.undo()).toBe(true);
		expect(getCell(sheet, 8, 2)?.value).toBe(600);
		expect(sheet.rowInfo.get(3)?.outlineLevel).toBe(1);

		expect(session.redo()).toBe(true);
		expect(getCell(sheet, 6, 0)).toBeUndefined();
		expect(sheet.rowInfo.get(1)?.outlineLevel).toBeUndefined();
	});

	it('supports hasHeader: false and pageBreakBetweenGroups', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		session.setRangeValues(0, { row: 0, col: 0 }, [
			['A', 10],
			['A', 20],
			['B', 30],
		]);
		const res = session.applySubtotal(0, parseRange('A1:B3')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [1],
			hasHeader: false,
			pageBreakBetweenGroups: true,
		});
		const sheet = workbook.sheets[0]!;
		expect(res.subtotalRows).toEqual([2, 4]);
		expect(res.grandTotalRow).toBe(5);
		expect(getCell(sheet, 2, 1)?.value).toBe(30);
		expect(getCell(sheet, 4, 1)?.value).toBe(30);
		expect(getCell(sheet, 5, 1)?.value).toBe(60);
		expect(sheet.preserved.get('rowBreaks')?.[0]).toContain('id="3"');
	});

	it('updates formulas on other sheets when subtotal rows are inserted and undone', () => {
		const { session, workbook } = setupSales();
		const s2 = session.addSheet('Summary');
		// Reference row 6 (Eve, 100) before subtotals are inserted
		session.setCellInput(s2, 0, 0, '=Sheet1!C6');
		expect(getCell(workbook.sheets[s2]!, 0, 0)?.value).toBe(100);

		session.applySubtotal(0, parseRange('A1:D6')!, {
			atEachChangeIn: 0,
			fn: 'sum',
			addSubtotalTo: [2],
		});
		// Row 6 shifted down by 1 (East Total inserted at row 4) to C7
		expect(getCell(workbook.sheets[s2]!, 0, 0)?.formula).toBe('Sheet1!C7');
		expect(getCell(workbook.sheets[s2]!, 0, 0)?.value).toBe(100);

		session.undo();
		expect(getCell(workbook.sheets[s2]!, 0, 0)?.formula).toBe('Sheet1!C6');
		expect(getCell(workbook.sheets[s2]!, 0, 0)?.value).toBe(100);
	});

	it('validates range and options', () => {
		const { session } = setupSales();
		expect(() =>
			session.applySubtotal(0, parseRange('A1:D6')!, {
				atEachChangeIn: 5,
				fn: 'sum',
				addSubtotalTo: [2],
			}),
		).toThrow(RangeError);
		expect(() =>
			session.applySubtotal(0, parseRange('A1:D6')!, {
				atEachChangeIn: 0,
				fn: 'sum',
				addSubtotalTo: [],
			}),
		).toThrow(RangeError);
		session.merge(0, parseRange('A2:B2')!, 'merge');
		expect(() =>
			session.applySubtotal(0, parseRange('A1:D6')!, {
				atEachChangeIn: 0,
				fn: 'sum',
				addSubtotalTo: [2],
			}),
		).toThrow(/merged/i);
	});
});
