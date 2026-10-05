// Iterative calculation (`calcPr iterate`): circular references converge instead of being
// reported and left at 0.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/load.js';
import { saveXlsx } from '../write/save.js';
import { CIRCULAR_REFERENCE_WARNING } from './engine.js';
import { book, engine, get } from './test-helpers.js';

const iterating = (cells: Parameters<typeof book>[0], count = 100, delta = 0.001) => {
	const wb = book(cells);
	wb.iterate = { count, delta };
	return wb;
};

describe('iterative calculation', () => {
	it('leaves circular references alone when it is off', () => {
		const wb = book({ A1: '=A1+1' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(0);
		expect(wb.warnings).toContain(CIRCULAR_REFERENCE_WARNING);
	});

	it('runs a self-reference for the iteration count', () => {
		const wb = iterating({ A1: '=A1+1' }, 25);
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(25);
		expect(wb.warnings).not.toContain(CIRCULAR_REFERENCE_WARNING);
	});

	it('stops once a loop settles within the maximum change', () => {
		// A1 and B1 each average the other with 10, so both approach 10.
		const wb = iterating({ A1: '=B1/2+5', B1: '=A1/2+5' }, 100, 0.0001);
		const calc = engine(wb);
		calc.recalculateAll();
		expect(get(wb, 'A1') as number).toBeCloseTo(10, 3);
		expect(get(wb, 'B1') as number).toBeCloseTo(10, 3);
		expect(calc.circularCells()).toHaveLength(0);
	});

	it('uses fewer passes when the maximum change is looser', () => {
		const tight = iterating({ A1: '=A1/2+5' }, 100, 1e-9);
		const loose = iterating({ A1: '=A1/2+5' }, 100, 1);
		engine(tight).recalculateAll();
		engine(loose).recalculateAll();
		expect(get(tight, 'A1') as number).toBeCloseTo(10, 6);
		expect(Math.abs((get(loose, 'A1') as number) - 10)).toBeGreaterThan(0.1);
	});

	it('feeds the converged value to formulas that depend on the loop', () => {
		const wb = iterating({ A1: '=B1/2+5', B1: '=A1/2+5', C1: '=A1*2' }, 100, 1e-9);
		engine(wb).recalculateAll();
		expect(get(wb, 'C1') as number).toBeCloseTo(20, 5);
	});

	it('lets a loop read a fixed input, as in a running balance', () => {
		// balance = balance * 1.1 + deposit, evaluated for 3 passes from 0.
		const wb = iterating({ B1: 100, A1: '=A1*1.1+B1' }, 3, 0);
		engine(wb).recalculateAll();
		expect(get(wb, 'A1') as number).toBeCloseTo(100 + 110 + 121, 8);
	});
});

describe('iterative calculation in the file', () => {
	it('is read from calcPr and written back', async () => {
		const wb = book({});
		wb.iterate = { count: 50, delta: 0.01 };
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.iterate).toEqual({ count: 50, delta: 0.01 });
		delete back.iterate;
		expect((await loadXlsx(await saveXlsx(back))).iterate).toBeUndefined();
	});

	it('defaults to 100 passes and a 0.001 change when the counts are absent', async () => {
		const wb = book({ A1: 1 });
		wb.iterate = { count: 7, delta: 5 };
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const xml = await zip.file('xl/workbook.xml')!.async('string');
		expect(xml).toContain('iterate="1" iterateCount="7" iterateDelta="5"');
		zip.file('xl/workbook.xml', xml.replace(' iterateCount="7" iterateDelta="5"', ''));
		const back = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(back.iterate).toEqual({ count: 100, delta: 0.001 });
	});

	it('keeps a file that turns iteration off from turning it on', async () => {
		const back = await loadXlsx(await saveXlsx(book({ A1: 1 })));
		expect(back.iterate).toBeUndefined();
	});
});
