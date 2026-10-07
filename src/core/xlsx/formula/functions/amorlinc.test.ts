import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { calc, book, engine, get, set, calcArray } from '../test-helpers.js';
import { FUNCTION_CATALOG } from './registry.js';
import { loadXlsx } from '../../read/index.js';
import { saveXlsx } from '../../write/index.js';
import { getCell } from '../../cells.js';

// Native Excel 16 results, recorded by scripts/record-xlsx-amorlinc.ps1.
const fixture = JSON.parse(
	readFileSync(new URL('../__fixtures__/amorlinc-excel.json', import.meta.url), 'utf8'),
) as {
	version: string;
	cases: { formula: string; type: string; value: number | string }[];
};
describe('AMORLINC native Excel compatibility', () => {
	it.each(fixture.cases)('$formula', ({ formula, type, value }) => {
		const actual = calc(formula.slice(1));
		if (type === 'error') expect(actual).toEqual({ error: value });
		else {
			expect(typeof actual).toBe('number');
			expect(Math.abs((actual as number) - (value as number))).toBeLessThan(1e-9);
		}
	});
	it('is discoverable in the shared function catalog', () => {
		expect(FUNCTION_CATALOG.find((fn) => fn.name === 'AMORLINC')?.category).toBe('Financial');
	});
	it('recalculates when referenced inputs change', () => {
		const wb = book({ A1: 2400, B1: '=AMORLINC(A1,39679,39813,300,1,0.15)' });
		const calculator = engine(wb);
		calculator.recalculateAll();
		expect(get(wb, 'B1')).toBe(360);
		set(wb, 'A1', 4000);
		calculator.recalculateFrom([{ sheet: 0, row: 0, col: 0 }]);
		expect(get(wb, 'B1')).toBe(600);
	});
	it('lifts period arrays using the normal dynamic-array engine', () => {
		expect(calcArray('AMORLINC(2400,39679,39813,300,{1,7},0.15)', {}, 1, 2)).toEqual([[360, 0]]);
	});
	it.each([false, true])(
		'survives save/reload and recalculation with date1904=%s',
		async (date1904) => {
			const formula = 'AMORLINC(2400,DATE(2020,2,29),DATE(2020,3,31),300,A1,0.15,1)';
			const wb = book({ A1: 0, B1: `=${formula}` }, ['Sheet1']);
			wb.date1904 = date1904;
			engine(wb).recalculateAll();
			expect(get(wb, 'B1')).toBeCloseTo(31.475409836065573, 10);
			const loaded = await loadXlsx(await saveXlsx(wb));
			expect(loaded.date1904).toBe(date1904);
			expect(getCell(loaded.sheets[0]!, 0, 1)?.formula).toBe(formula);
			expect(get(loaded, 'B1')).toBeCloseTo(31.475409836065573, 10);
			set(loaded, 'A1', 1);
			engine(loaded).recalculateAll();
			expect(get(loaded, 'B1')).toBe(360);
		},
	);
});
