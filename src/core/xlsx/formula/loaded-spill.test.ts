// A dynamic array read from a file arrives with its spilled results already in the cells. Until its
// anchor is evaluated, those cells are the anchor's spill: overwriting the anchor or shrinking its
// result must clear them and recalculate the formulas that read them, as Excel does.
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { createEditSession } from '../edit/index';
import type { Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { isSpilledCell, type SpilledCell } from './spill';

/** A1 `=SEQUENCE(B1)` saved with its three results in A1:A3, and C3 `=A3*10` reading the last. */
async function loadedSpill(): Promise<Workbook> {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	putCell(sheet, 0, 1, { value: 3 });
	putCell(sheet, 0, 0, { value: 1, formula: 'SEQUENCE(B1)', dynamicArray: true });
	const spilled = (value: number): SpilledCell => ({ value, spillAnchor: { row: 0, col: 0 } });
	putCell(sheet, 1, 0, spilled(2));
	putCell(sheet, 2, 0, spilled(3));
	putCell(sheet, 2, 2, { value: 30, formula: 'A3*10' });
	const loaded = await loadXlsx(await saveXlsx(wb));
	expect(isSpilledCell(getCell(loaded.sheets[0]!, 2, 0))).toBe(true);
	return loaded;
}

const values = (wb: Workbook) =>
	['A1', 'A2', 'A3', 'C3'].map((ref) => {
		const row = Number(ref.slice(1)) - 1;
		const col = ref.charCodeAt(0) - 65;
		return getCell(wb.sheets[0]!, row, col)?.value ?? null;
	});

describe.each([
	['before the graph is prepared', false],
	['after prepareCalculation', true],
])('a loaded spill, %s', (_, prepare) => {
	async function session() {
		const wb = await loadedSpill();
		const edit = createEditSession(wb, { autoRowHeight: false });
		if (prepare) edit.prepareCalculation();
		return { wb, edit };
	}

	it('is cleared when its anchor is overwritten, and its readers recalculate', async () => {
		const { wb, edit } = await session();
		edit.setCellValue(0, 0, 0, 5);
		expect(values(wb)).toEqual([5, null, null, 0]);
		expect(isSpilledCell(getCell(wb.sheets[0]!, 2, 0))).toBe(false);
	});

	it('loses the cells its result no longer covers', async () => {
		const { wb, edit } = await session();
		edit.setCellValue(0, 0, 1, 1);
		expect(values(wb)).toEqual([1, null, null, 0]);
	});

	it('keeps spilling over its cells when the result keeps its size', async () => {
		const { wb, edit } = await session();
		edit.setCellInput(0, 0, 0, '=SEQUENCE(B1,1,10)');
		expect(values(wb)).toEqual([10, 11, 12, 120]);
		expect(isSpilledCell(getCell(wb.sheets[0]!, 2, 0))).toBe(true);
	});
});
