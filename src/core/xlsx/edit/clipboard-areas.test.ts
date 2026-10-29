import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};

function sheetWith() {
	const wb = createWorkbook();
	const s = createEditSession(wb, { recalc: false });
	const rows: number[][] = [];
	for (let r = 1; r <= 4; r++) rows.push([1, 2, 3, 4, 5, 6].map((c) => r * 10 + c));
	s.setRangeValues(0, { row: 0, col: 0 }, rows);
	const at = (row: number, col: number) => getCell(wb.sheets[0]!, row, col)?.value ?? null;
	return { wb, s, at };
}

describe('multi-area copy (checked against Excel 16)', () => {
	it('pastes areas with the same rows side by side, gaps dropped', () => {
		const { s, at } = sheetWith();
		const payload = s.copyAreas(0, [R('C1:C2'), R('A1:A2')]);
		expect(payload.tsv).toBe('11\t13\r\n21\t23\r\n');
		s.paste(0, { row: 0, col: 7 }, payload);
		expect([at(0, 7), at(0, 8), at(1, 7), at(1, 8)]).toEqual([11, 13, 21, 23]);
		expect(at(0, 9)).toBeNull();
	});

	it('pastes areas with the same columns stacked', () => {
		const { s, at } = sheetWith();
		s.paste(0, { row: 5, col: 7 }, s.copyAreas(0, [R('A1:B1'), R('A3:B3')]));
		expect([at(5, 7), at(5, 8), at(6, 7), at(6, 8)]).toEqual([11, 12, 31, 32]);
	});

	it('refuses other shapes with the message Excel gives', () => {
		const { s } = sheetWith();
		for (const shape of [
			['A1:A2', 'C2:C3'],
			['A1:B2', 'D3:E4'],
			['A1:A2', 'C1:C3'],
			['A1:A2', 'A4:B4'],
		])
			expect(() => s.copyAreas(0, shape.map(R))).toThrow(
				"This action won't work on multiple selections.",
			);
	});

	it('keeps relative formulas per cell when the gap is dropped', () => {
		const { wb, s } = sheetWith();
		s.setCellInput(0, 0, 8, '=I2');
		s.setCellInput(0, 0, 10, '=K2');
		s.paste(0, { row: 0, col: 13 }, s.copyAreas(0, [R('K1'), R('I1')]));
		const sheet = wb.sheets[0]!;
		expect(getCell(sheet, 0, 13)?.formula).toBe('N2');
		expect(getCell(sheet, 0, 14)?.formula).toBe('O2');
	});
});

describe('large selections', () => {
	it('tiles a copied cell over a million-cell selection', () => {
		const { s, at } = sheetWith();
		const payload = s.copy(0, R('A1'));
		const started = Date.now();
		s.paste(0, R('A10:AMJ1009'), payload);
		expect(at(9, 0)).toBe(11);
		expect(at(1008, 1023)).toBe(11);
		expect(Date.now() - started).toBeLessThan(30_000);
		s.undo();
		expect(at(9, 0)).toBeNull();
	});
});
