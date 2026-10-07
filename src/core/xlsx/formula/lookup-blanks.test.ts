import { describe, expect, it } from 'vitest';
import { calc, E } from './test-helpers.js';
import { findExact } from './functions/lookup-core.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from '../edit/index.js';
import { getCell } from '../cells.js';
import { saveXlsx } from '../write/index.js';
import { loadXlsx } from '../read/index.js';

// Scalar results independently checked in Excel 16.0 build 20430, 2026-10-07.
// A blank lookup reference differs from an explicit "" string or zero.
describe('Excel lookup of trailing blank cells', () => {
	it.each([
		['XMATCH(A2,A1:A5)', 2],
		['XMATCH(A2,A1:A5,0,-1)', 5],
		['XMATCH("",A1:A5)', E.NA],
		['XMATCH(0,A1:A5)', E.NA],
		['MATCH(0,A1:A5,0)', E.NA],
		['MATCH("",A1:A5,0)', E.NA],
	] as const)('%s matches recorded Excel behavior', (formula, result) => {
		expect(calc(formula, { A1: 1, B1: 'one' })).toEqual(result);
	});

	it('finds blanks in references entirely beyond the used rows and columns', () => {
		expect(calc('XMATCH(A2,A10:A15)', { A1: 1 })).toBe(1);
		expect(calc('XMATCH(A2,A10:A15,0,-1)', { A1: 1 })).toBe(6);
		expect(calc('XMATCH(A2,AB1:AF1)', { A1: 1 })).toBe(1);
		expect(calc('XMATCH(A2,AB1:AF1,0,-1)', { A1: 1 })).toBe(5);
	});

	it('searches whole-column blank tails without reading a million cells', () => {
		let reads = 0;
		const vector = {
			length: 1,
			logicalLength: 1_048_576,
			get: () => {
				reads++;
				return 1;
			},
		};
		expect(findExact(vector, null, false)).toBe(1);
		expect(findExact(vector, null, false, true)).toBe(1_048_575);
		expect(reads).toBe(1);
		expect(calc('XMATCH(A2,B:B,0,-1)', { A1: 1 })).toBe(1_048_576);
	});

	it('distinguishes stored empty strings from blank cells', () => {
		expect(calc('XMATCH(A2,B1:B5)', { B1: '=""' })).toBe(2);
		expect(calc('XMATCH(A2,B1:B5,0,-1)', { B1: '=""' })).toBe(5);
		expect(calc('XMATCH(A2,B1)', { B1: '=""' })).toEqual(E.NA);
		expect(calc('XMATCH(A2,{""})')).toEqual(E.NA);
		expect(calc('XMATCH("",B1:B5)', { B1: '=""' })).toBe(1);
	});

	it('preserves blank XLOOKUP results and formulas through save and reopen', async () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		session.setCellValue(0, 0, 0, 1);
		session.setCellInput(0, 0, 25, '=XLOOKUP(A2,A1:A5,B1:B5,"missing")');
		expect(getCell(workbook.sheets[0]!, 0, 25)?.value).toBe(0);
		const loaded = await loadXlsx(await saveXlsx(workbook));
		expect(getCell(loaded.sheets[0]!, 0, 25)).toMatchObject({
			formula: 'XLOOKUP(A2,A1:A5,B1:B5,"missing")',
			value: 0,
		});
	});
});
