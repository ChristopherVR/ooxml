import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address';
import { getCell, putCell } from '../cells';
import type { CellValue, Workbook } from '../model';
import { internStyle } from '../styles';
import { loadXlsx } from '../read/load';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/save';
import { createEditSession } from './session';

const R = (ref: string) => parseRange(ref)!;
const A = (ref: string) => parseAddress(ref)!;
const col = (wb: Workbook, c: number, from: number, to: number): CellValue[] =>
	Array.from(
		{ length: to - from + 1 },
		(_v, i) => getCell(wb.sheets[0]!, from + i, c)?.value ?? null,
	);

function seeded() {
	const wb = createWorkbook();
	const s = createEditSession(wb, { recalc: false });
	s.setRangeValues(0, A('A1'), [['Item'], ['b'], ['c'], ['a']]);
	return { wb, s };
}

describe('table sort state', () => {
	it('writes the sort into the table part and reads it back', async () => {
		const { wb, s } = seeded();
		s.createTable(0, R('A1:A4'), true);
		s.sort(0, R('A1:A4'), [{ col: 0, descending: true }], true);
		const table = wb.sheets[0]!.tables[0]!;
		expect(table.sortState?.conditions[0]?.descending).toBe(true);
		expect(wb.sheets[0]!.sortState).toBeUndefined();
		const loaded = await loadXlsx(await saveXlsx(wb));
		const state = loaded.sheets[0]!.tables[0]!.sortState;
		expect(state?.range).toEqual(R('A2:A4'));
		expect(state?.conditions[0]).toMatchObject({ descending: true, ref: R('A2:A4') });
		expect(loaded.sheets[0]!.sortState).toBeUndefined();
		// A second save patches the source part and keeps the sort.
		const again = await loadXlsx(await saveXlsx(loaded));
		expect(again.sheets[0]!.tables[0]!.sortState).toEqual(state);
	});

	it('replaces a source sort when the table is sorted again', async () => {
		const { wb, s } = seeded();
		s.createTable(0, R('A1:A4'), true);
		s.sort(0, R('A1:A4'), [{ col: 0, descending: true }], true);
		const loaded = await loadXlsx(await saveXlsx(wb));
		createEditSession(loaded, { recalc: false }).sort(0, R('A1:A4'), [{ col: 0 }], true);
		const again = await loadXlsx(await saveXlsx(loaded));
		expect(again.sheets[0]!.tables[0]!.sortState?.conditions[0]?.descending).toBeUndefined();
	});
});

describe('reapply', () => {
	it('re-runs a sheet sort after the data changed', () => {
		const { wb, s } = seeded();
		s.sort(0, R('A1:A4'), [{ col: 0 }], true);
		expect(col(wb, 0, 1, 3)).toEqual(['a', 'b', 'c']);
		s.setCellInput(0, 1, 0, 'z');
		s.reapplyFilter(0, A('A2'));
		expect(col(wb, 0, 1, 3)).toEqual(['b', 'c', 'z']);
	});

	it('re-runs a table sort for a cell inside the table', () => {
		const { wb, s } = seeded();
		s.createTable(0, R('A1:A4'), true);
		s.sort(0, R('A1:A4'), [{ col: 0, descending: true }], true);
		s.setCellInput(0, 3, 0, 'zz');
		s.reapplyFilter(0, A('A2'));
		expect(col(wb, 0, 1, 3)).toEqual(['zz', 'c', 'b']);
	});

	it('re-runs a colour sort and the filter criteria together', () => {
		const wb = createWorkbook();
		const red = internStyle(wb, {
			...wb.styles[0]!,
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } },
		});
		const sheet = wb.sheets[0]!;
		['H', 'a', 'b', 'c'].forEach((v, i) => putCell(sheet, i, 0, { value: v }));
		const s = createEditSession(wb, { recalc: false });
		putCell(sheet, 3, 0, { value: 'c', styleId: red });
		s.sort(0, R('A1:A4'), [{ col: 0, sortOn: 'cellColor', color: '#FF0000' }], true);
		expect(col(wb, 0, 1, 3)).toEqual(['c', 'a', 'b']);
		putCell(sheet, 3, 0, { value: 'b', styleId: red });
		s.reapplyFilter(0, A('A2'));
		expect(col(wb, 0, 1, 3)).toEqual(['c', 'b', 'a']);
	});
});
