import { describe, expect, it } from 'vitest';
import { analyzeVisioMasterFormula } from './formula-master.js';
import { analyzeVisioFormula } from './formula.js';

const local = (formula: string) => analyzeVisioMasterFormula(formula, { textFieldFree: true });
describe('bounded master-context formula locality', () => {
	it('admits explicit redirect reads only for a single direct cell reference', () => {
		expect(local('SETATREF(Controls.visSSTXT.Y)')).toMatchObject({
			dynamic: false,
			references: [{ cell: 'Controls.visSSTXT.Y' }],
		});
		expect(analyzeVisioFormula('SETATREF(Controls.visSSTXT.Y)').dynamic).toBe(true);
		for (const formula of [
			'SETATREF("Width")',
			'SETATREF(Width+1)',
			'SETATREF(Width,2)',
			'SETATREFEXPR(Width)',
		])
			expect(local(formula).dynamic, formula).toBe(true);
	});
	it('requires actual field-free local text for measurement and text extraction', () => {
		for (const formula of [
			'TEXTWIDTH(TheText)',
			'TEXTHEIGHT(TheText,TxtWidth)',
			'SHAPETEXT(TheText)',
		]) {
			expect(local(formula)).toMatchObject({ dynamic: false, readsText: true });
			expect(analyzeVisioMasterFormula(formula, { textFieldFree: false }).dynamic).toBe(true);
		}
		expect(local('TEXTHEIGHT(TheText,TxtWidth)').references).toContainEqual({ cell: 'TxtWidth' });
		for (const formula of [
			'TEXTWIDTH("TheText")',
			'TEXTWIDTH(Sheet.1!TheText)',
			'SHAPETEXT(TheText,1)',
			'TEXTHEIGHT(TheText)',
		])
			expect(local(formula).dynamic, formula).toBe(true);
	});
	it('records explicit and implicit color reads without evaluating colors or text', () => {
		expect(local('LUM(THEMEVAL("BackgroundColor"))')).toMatchObject({
			dynamic: false,
			readsTheme: true,
		});
		expect(local('LUM("FillForegnd")').references).toContainEqual({ cell: 'FillForegnd' });
		expect(local('LUM("Pages[Page 1]!Sheet.1!FillForegnd")').dynamic).toBe(true);
		expect(local('IF(STRSAME(SHAPETEXT(TheText),""),Width,Height)')).toMatchObject({
			dynamic: false,
			readsText: true,
		});
	});
	it('keeps opaque or cross-page lookup functions unsafe and bounded', () => {
		for (const formula of [
			'EVALCELL("Sheet.1!Width")',
			'UNKNOWN(Width)',
			'TEXTWIDTH(EVALTEXT("TheText"))',
		])
			expect(local(formula).dynamic, formula).toBe(true);
		expect(() => local('1'.repeat(8193))).toThrow(/length limit/);
		expect(() => local('('.repeat(70) + 'Width' + ')'.repeat(70))).toThrow(/depth limit/);
	});
});
