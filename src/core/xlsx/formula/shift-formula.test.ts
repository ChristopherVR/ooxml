import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/load';
import { forgetFormulaTemplates, renameSheetInFormula, shiftFormula } from './shift-formula';

const FORMULAS = [
	'A1+B2',
	'=SUM($A$1:A10)*2',
	'SUM(A3 : A6)+A1 : B2',
	"Sheet2!A5+'Other Sheet'!B$7+[1]Sheet2!A5",
	'SUM(Sheet2:Sheet3!A1:A4)',
	'A:A+1:3+$B:$C+$2:$2',
	'SEQUENCE(3)+A2#+Sheet2!C3#',
	'"A1 is text"&A1&LOG10(B2)',
	'OFFSET(A1,2,0)+INDEX(B:B,4):INDEX(B:B,9)',
	'Sheet2!Total+Table1[Amount]+#REF!+Sheet2!#REF!',
	"'O''Brien'!A3*2",
	'XFD1048576+XFC1048575',
	'MAX(AA550-MAX((IF(AA550>0,MIN(Dashboard!$B$8,AA550+AC551),0))),0)',
	'B5:A1',
	'{1,2;3,4}+A1',
];

type Step = (formula: string) => string;
const STEPS: Step[] = [
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'Sheet1', axis: 'row', at: 0, count: 1 }),
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'Sheet2', axis: 'row', at: 3, count: 2 }),
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'sheet1', axis: 'row', at: 2, count: -3 }),
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'Sheet1', axis: 'col', at: 1, count: 1 }),
	(f) => renameSheetInFormula(f, 'Sheet2', 'My Data'),
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'My Data', axis: 'col', at: 0, count: -1 }),
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'Sheet1', axis: 'row', at: 1_048_570, count: 3 }),
	(f) => renameSheetInFormula(f, "O'Brien", 'Sheet2'),
	(f) => shiftFormula(f, 'Sheet1', { sheet: 'Sheet1', axis: 'col', at: 16_380, count: 2 }),
	(f) => renameSheetInFormula(f, 'my data', 'Sheet2'),
];

/** Each formula through every step, with templates cached across steps or tokenized afresh. */
function run(formulas: string[], fresh: boolean): string[][] {
	forgetFormulaTemplates();
	return formulas.map((formula) => {
		const out: string[] = [];
		let text = formula;
		for (const step of STEPS) {
			if (fresh) forgetFormulaTemplates();
			text = step(text);
			out.push(text);
		}
		return out;
	});
}

describe('cached formula templates', () => {
	it('rewrite exactly like tokenizing every formula again', async () => {
		const workbook = await loadXlsx(
			new Uint8Array(
				readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-features.xlsx')),
			),
		);
		const formulas = [...FORMULAS];
		for (const sheet of workbook.sheets)
			for (const cells of sheet.rows.values())
				for (const cell of cells.values()) if (cell.formula) formulas.push(cell.formula);
		expect(run(formulas, false)).toEqual(run(formulas, true));
	});

	it('derive the rewritten text', () => {
		forgetFormulaTemplates();
		const spec = { sheet: 'Sheet1', axis: 'row' as const, at: 2, count: 1 };
		const once = shiftFormula('A1+A3 : A5+Sheet2!A3', 'Sheet1', spec);
		expect(once).toBe('A1+A4:A6+Sheet2!A3');
		expect(shiftFormula(once, 'Sheet1', spec)).toBe('A1+A5:A7+Sheet2!A3');
		expect(renameSheetInFormula('Sheet2!A3+A1', 'Sheet2', 'Q1 Data')).toBe("'Q1 Data'!A3+A1");
	});
});
