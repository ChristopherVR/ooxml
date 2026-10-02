import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/index.js';
import { createEditSession } from './session.js';

const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const at = (wb: Workbook, sheet: number, row: number, col: number) =>
	getCell(wb.sheets[sheet] ?? wb.sheets[0]!, row, col);

function setup() {
	const wb = createWorkbook({ sheets: ['Data', 'Report'] });
	const s = createEditSession(wb, { recalc: true });
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Item', 'Amount'],
		['a', 1],
		['b', 2],
		['c', 3],
	]);
	s.createTable(0, R('A1:B4'), true);
	s.setCellInput(1, 0, 0, '=SUM(Table1[Amount])');
	s.setCellInput(0, 1, 2, '=Table1[@Amount]*2');
	return { wb, s, table: () => wb.sheets[0]!.tables[0]! };
}

describe('table edits', () => {
	it('renames a table and its structured references', () => {
		const { wb, s, table } = setup();
		expect(() => s.updateTable(0, 0, { name: 'A1' })).toThrow();
		s.updateTable(0, 'table1', {
			name: 'Sales',
			showRowStripes: false,
			styleName: 'TableStyleLight9',
		});
		expect(table()).toMatchObject({ name: 'Sales', displayName: 'Sales', showRowStripes: false });
		expect(at(wb, 1, 0, 0)?.formula).toBe('SUM(Sales[Amount])');
		expect(at(wb, 1, 0, 0)?.value).toBe(6);
		expect(at(wb, 0, 1, 2)?.formula).toBe('Sales[@Amount]*2');
		expect(s.undoLabel()).toBe('Rename table');
		s.undo();
		expect(table().name).toBe('Table1');
		expect(at(wb, 1, 0, 0)?.formula).toBe('SUM(Table1[Amount])');
	});

	it('toggles the totals and header rows', async () => {
		const { wb, s, table } = setup();
		s.updateTable(0, 0, { totalsRow: true });
		expect(table().range).toEqual(R('A1:B5'));
		expect(at(wb, 0, 4, 0)?.value).toBe('Total');
		expect(at(wb, 0, 4, 1)?.formula).toBe('SUBTOTAL(109,Table1[Amount])');
		expect(at(wb, 0, 4, 1)?.value).toBe(6);
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets[0]?.tables[0]).toMatchObject({ totalsRow: true, range: R('A1:B5') });
		s.updateTable(0, 0, { totalsRow: false });
		expect(table().range).toEqual(R('A1:B4'));
		expect(at(wb, 0, 4, 1)).toBeUndefined();
		s.updateTable(0, 0, { headerRow: false });
		expect(table().range).toEqual(R('A2:B4'));
		expect(at(wb, 0, 0, 0)).toBeUndefined();
		// The row above is free again, so the header comes back in place.
		s.updateTable(0, 0, { headerRow: true });
		expect(table().range).toEqual(R('A1:B4'));
		expect(at(wb, 0, 0, 1)?.value).toBe('Amount');
	});

	it('shifts cells down when a new header row has no room', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setRangeValues(0, { row: 0, col: 0 }, [['x'], ['H'], [1], [2]]);
		s.createTable(0, R('A2:A4'), true);
		s.updateTable(0, 0, { headerRow: false });
		s.setCellValue(0, 1, 0, 'blocker');
		s.updateTable(0, 0, { headerRow: true });
		const table = wb.sheets[0]!.tables[0]!;
		expect(table.range).toEqual(R('A3:A5'));
		expect(at(wb, 0, 1, 0)?.value).toBe('blocker');
		expect(at(wb, 0, 2, 0)?.value).toBe('H');
		expect(at(wb, 0, 3, 0)?.value).toBe(1);
	});

	it('resizes a table, naming new columns from their headers', () => {
		const { s, table } = setup();
		s.setCellValue(0, 0, 2, 'Double');
		s.resizeTable(0, 0, R('A1:C6'));
		expect(table().range).toEqual(R('A1:C6'));
		expect(table().columns.map((c) => c.name)).toEqual(['Item', 'Amount', 'Double']);
		expect(() => s.resizeTable(0, 0, R('A2:C6'))).toThrow(/same row/);
		s.undo();
		expect(table().range).toEqual(R('A1:B4'));
	});

	it('converts a table to a range with absolute references', async () => {
		const { wb, s } = setup();
		s.convertTableToRange(0, 'Table1');
		expect(wb.sheets[0]?.tables).toEqual([]);
		expect(at(wb, 1, 0, 0)?.formula).toBe('SUM(Data!$B$2:$B$4)');
		expect(at(wb, 1, 0, 0)?.value).toBe(6);
		expect(at(wb, 0, 1, 2)?.formula).toBe('$B2*2');
		expect(s.undoLabel()).toBe('Convert to range');
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets[0]?.tables).toEqual([]);
		s.undo();
		expect(wb.sheets[0]?.tables.length).toBe(1);
	});
});
