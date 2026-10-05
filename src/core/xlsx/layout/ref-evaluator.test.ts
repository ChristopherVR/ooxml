import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import { createCalcEngine } from '../formula/index.js';
import { createWorkbook, createWorksheet } from '../workbook.js';
import { createRefEvaluator } from './ref-evaluator.js';

function book() {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const other = createWorksheet('My Sheet', 2);
	wb.sheets.push(other);
	for (let r = 0; r < 3; r++) {
		putCell(sheet, r, 0, { value: r + 1 });
		putCell(sheet, r, 1, { value: `x${r}` });
		putCell(other, r, 2, { value: (r + 1) * 10 });
	}
	wb.definedNames.push({ name: 'Vals', formula: 'Sheet1!$A$1:$A$3' });
	return wb;
}

describe('createRefEvaluator', () => {
	it('reads plain and sheet-qualified references row-major', () => {
		const wb = book();
		const evaluate = createRefEvaluator(wb);
		expect(evaluate('Sheet1!$A$1:$B$2')).toEqual([1, 'x0', 2, 'x1']);
		expect(evaluate("='My Sheet'!$C$1:$C$3")).toEqual([10, 20, 30]);
		expect(evaluate('A2:A3')).toEqual([2, 3]);
	});

	it('clips whole columns to the used area', () => {
		expect(createRefEvaluator(book())('Sheet1!A:A')).toEqual([1, 2, 3]);
	});

	it('resolves names and expressions through the calc engine', () => {
		const wb = book();
		const calc = createCalcEngine(wb);
		expect(createRefEvaluator(wb, calc)('Vals')).toEqual([1, 2, 3]);
		expect(createRefEvaluator(wb, calc)('OFFSET(Sheet1!A1,1,0,2,1)')).toEqual([2, 3]);
		expect(createRefEvaluator(wb)('Vals')).toEqual([]);
	});

	it('returns nothing for unknown sheets or junk', () => {
		const evaluate = createRefEvaluator(book());
		expect(evaluate('Nope!A1:A2')).toEqual([]);
		expect(evaluate('')).toEqual([]);
		expect(evaluate('SUM(')).toEqual([]);
	});
});
