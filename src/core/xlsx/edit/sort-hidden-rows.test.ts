import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import type { Workbook } from '../model';
import { loadXlsx } from '../read/load';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/save';
import { createEditSession } from './session';

/** Parses a test range, including its original row gaps. */
const range = (ref: string) => parseRange(ref)!;

/** Creates filtered records and both SUBTOTAL variants outside the sorted area. */
function setup() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Kind', 'Amount'],
		['keep', 10],
		['drop', 50],
		['keep', 30],
	]);
	session.setCellInput(0, 5, 1, '=SUBTOTAL(9,B2:B4)');
	session.setCellInput(0, 6, 1, '=SUBTOTAL(109,B2:B4)');
	session.setAutoFilter(0, range('A1:B4'));
	session.filterColumn(0, 0, ['keep']);
	return { workbook, session, sheet: workbook.sheets[0]! };
}

/** Checks the records, visibility and recalculated subtotals after a descending sort. */
function expectFilteredSort(workbook: Workbook) {
	const sheet = workbook.sheets[0]!;
	expect([1, 2, 3].map((r) => getCell(sheet, r, 0)?.value)).toEqual(['keep', 'drop', 'keep']);
	expect([1, 2, 3].map((r) => getCell(sheet, r, 1)?.value)).toEqual([30, 50, 10]);
	expect([1, 2, 3].map((r) => !!sheet.rowInfo.get(r)?.hidden)).toEqual([false, true, false]);
	expect(getCell(sheet, 5, 1)?.value).toBe(40);
	expect(getCell(sheet, 6, 1)?.value).toBe(40);
}

describe('sorting hidden rows', () => {
	it.each(['column', 'range'] as const)(
		'keeps filtered rows in place via %s, undo, redo and save',
		async (mode) => {
			const { workbook, session, sheet } = setup();
			const before = structuredClone(sheet);
			if (mode === 'column') session.sortByColumn(0, 1, true);
			else session.sort(0, range('A1:B4'), [{ col: 1, descending: true }], true);
			expectFilteredSort(workbook);
			expect(sheet.autoFilter).toEqual(before.autoFilter);
			const after = structuredClone(sheet);
			session.undo();
			expect(workbook.sheets[0]).toEqual(before);
			session.redo();
			expect(workbook.sheets[0]).toEqual(after);
			const loaded = await loadXlsx(await saveXlsx(workbook));
			createEditSession(loaded, { autoRowHeight: false }).calculateNow({ full: true });
			expectFilteredSort(loaded);
		},
	);

	it('keeps a loaded filter and its hidden record in place', async () => {
		const { workbook } = setup();
		const loaded = await loadXlsx(await saveXlsx(workbook));
		const session = createEditSession(loaded, { autoRowHeight: false });
		session.sortByColumn(0, 1, true);
		expectFilteredSort(loaded);
	});

	it('preserves opaque loaded criteria without reapplying unsupported filters', async () => {
		const { workbook, sheet } = setup();
		const sourceXml =
			'<filterColumn colId="0"><customFilters><customFilter operator="equal" val="keep"/></customFilters></filterColumn>';
		sheet.autoFilter!.columns = [{ offset: 0, sourceXml }];
		const loaded = await loadXlsx(await saveXlsx(workbook));
		const session = createEditSession(loaded, { autoRowHeight: false });
		const before = structuredClone(loaded.sheets[0]!.autoFilter);
		session.sortByColumn(0, 1, true);
		expect(loaded.sheets[0]!.autoFilter).toEqual(before);
		expect([1, 2, 3].map((r) => getCell(loaded.sheets[0]!, r, 1)?.value)).toEqual([30, 50, 10]);
		const saved = await loadXlsx(await saveXlsx(loaded));
		expect(saved.sheets[0]!.autoFilter?.columns?.[0]?.sourceXml).toContain('operator="equal"');
		expect(saved.sheets[0]!.rowInfo.get(2)?.hidden).toBe(true);
	});

	it('keeps manually hidden records and row dimensions in place without a filter', () => {
		const { workbook, session, sheet } = setup();
		session.setAutoFilter(0, undefined);
		session.setRowHeight(0, [2], 25);
		session.setHidden(0, 'row', [2], true);
		const dimensions = structuredClone(sheet.rowInfo);
		session.sort(0, range('A1:B4'), [{ col: 1, descending: true }], true);
		expect([1, 2, 3].map((r) => getCell(sheet, r, 1)?.value)).toEqual([30, 50, 10]);
		expect(sheet.rowInfo).toEqual(dimensions);
		expect(getCell(workbook.sheets[0]!, 6, 1)?.value).toBe(40);
	});

	it('moves formulas, comments and links between visible row gaps only', () => {
		const { session, sheet } = setup();
		for (const r of [1, 2, 3]) session.setCellInput(0, r, 2, `=B${r + 1}*2`);
		session.setComment(0, { row: 1, col: 2 }, 'low', 'a');
		session.setComment(0, { row: 2, col: 2 }, 'hidden', 'a');
		session.setHyperlink(0, range('C4'), { target: 'https://example.com/high' });
		session.setCellValue(0, 1, 3, 'outside');
		session.sort(0, range('A1:C4'), [{ col: 1, descending: true }], true);
		expect([1, 2, 3].map((r) => getCell(sheet, r, 2)?.formula)).toEqual(['B2*2', 'B3*2', 'B4*2']);
		expect([1, 2, 3].map((r) => getCell(sheet, r, 2)?.value)).toEqual([60, 100, 20]);
		expect(sheet.comments.find((c) => c.text === 'low')?.address).toEqual({ row: 3, col: 2 });
		expect(sheet.comments.find((c) => c.text === 'hidden')?.address).toEqual({ row: 2, col: 2 });
		expect(sheet.hyperlinks[0]?.range).toEqual(range('C2'));
		expect(getCell(sheet, 1, 3)?.value).toBe('outside');
	});

	it.each([{ hidden: [0, 3] }, { hidden: [0, 1, 3] }, { hidden: [0, 1, 2, 3] }])(
		'handles leading, trailing and all hidden rows: %j',
		({ hidden }) => {
			const workbook = createWorkbook();
			const session = createEditSession(workbook, { autoRowHeight: false });
			session.setRangeValues(0, { row: 0, col: 0 }, [[50], [3], [1], [60]]);
			session.setHidden(0, 'row', hidden, true);
			const expected = hidden.includes(1) ? [50, 3, 1, 60] : [50, 1, 3, 60];
			session.sort(0, range('A1:A4'), [{ col: 0 }], false);
			expect([0, 1, 2, 3].map((r) => getCell(workbook.sheets[0]!, r, 0)?.value)).toEqual(expected);
		},
	);

	it('keeps hidden table records and the totals row in place', () => {
		const { session, sheet } = setup();
		session.setAutoFilter(0, undefined);
		session.createTable(0, range('A1:B4'), true);
		session.updateTable(0, sheet.tables[0]!.name, { totalsRow: true });
		session.setHidden(0, 'row', [2], true);
		const totals = structuredClone(sheet.rows.get(4));
		session.sortByColumn(0, 1, true);
		expect([1, 2, 3].map((r) => getCell(sheet, r, 1)?.value)).toEqual([30, 50, 10]);
		expect(sheet.rows.get(4)).toEqual(totals);
	});
});
