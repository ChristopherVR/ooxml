import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { calc, calcArray, book, engine, get } from '../test-helpers';
import { loadXlsx } from '../../read/index';
import { saveXlsx } from '../../write/index';
import { getCell } from '../../cells';

// Excel 16.0 measurements, reproducible with scripts/record-xlsx-depreciation.ps1.
const fixture = JSON.parse(
	readFileSync(new URL('../__fixtures__/depreciation-excel.json', import.meta.url), 'utf8'),
) as {
	cases: { formula: string; type: string; value: number | string }[];
};
describe('native DB/DDB fractional-period compatibility', () => {
	it.each(fixture.cases)('$formula', ({ formula, type, value }) => {
		const actual = calc(formula.slice(1));
		if (type === 'error') expect(actual).toEqual({ error: value });
		else {
			expect(typeof actual).toBe('number');
			expect(Math.abs((actual as number) - (value as number))).toBeLessThan(
				1e-12 * Math.max(1, Math.abs(value as number)),
			);
		}
	});
	it('lifts fractional-period arrays without rounding to whole periods', () => {
		const values = calcArray('DDB(2400,300,10,{0.5,1.5,2.5})', {}, 1, 3)[0]!;
		expect(values[0]).toBe(480);
		expect(values[1] as number).toBeCloseTo(429.32505167996, 10);
		expect(values[2] as number).toBeCloseTo(343.460041343968, 10);
	});
	it('preserves fractional formulas and caches through save/reload', async () => {
		const wb = book({ A1: '=DDB(2400,300,10,1.5)', B1: '=DB(2400,300,10,1,7.9)' }, ['Sheet1']);
		engine(wb).recalculateAll();
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(getCell(loaded.sheets[0]!, 0, 0)?.formula).toBe('DDB(2400,300,10,1.5)');
		expect(getCell(loaded.sheets[0]!, 0, 1)?.formula).toBe('DB(2400,300,10,1,7.9)');
		expect(get(loaded, 'A1') as number).toBeCloseTo(429.32505167996, 10);
		expect(get(loaded, 'B1')).toBe(263.2);
		engine(loaded).recalculateAll();
		expect(get(loaded, 'A1') as number).toBeCloseTo(429.32505167996, 10);
		expect(get(loaded, 'B1')).toBe(263.2);
	});
});
