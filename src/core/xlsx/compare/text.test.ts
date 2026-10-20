import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange, type CellRange } from '../address';
import { putCell } from '../cells';
import { cellError, type Cell, type Workbook } from '../model';
import { internStyle, patchStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { compareWorkbooks } from './compare';
import { comparisonText } from './text';

const R = (ref: string): CellRange => {
	const range = parseRange(ref);
	if (!range) throw new Error(ref);
	return range;
};

function put(workbook: Workbook, ref: string, cell: Partial<Cell>, sheet = 0): void {
	const at = parseAddress(ref);
	if (!at) throw new Error(ref);
	putCell(workbook.sheets[sheet]!, at.row, at.col, { value: null, ...cell });
}

function pair(): [Workbook, Workbook] {
	const before = createWorkbook({ sheets: ['Data', 'Old', 'Notes'] });
	const after = createWorkbook({ sheets: ['Inputs', 'Notes', 'Extra'] });
	after.sheets[0]!.sheetId = 1;
	after.sheets[1]!.sheetId = 3;
	after.sheets[1]!.state = 'hidden';
	after.sheets[2]!.sheetId = 4;
	before.definedNames = [{ name: 'Rate', formula: 'Data!$A$1' }];
	after.definedNames = [
		{ name: 'Rate', formula: 'Inputs!$A$2' },
		{ name: 'Local', formula: '1', localSheet: 1, hidden: true },
	];
	after.sheets[0]!.merges = [R('D1:E1')];
	put(before, 'A1', { value: 1 });
	put(after, 'A1', { value: 'one' });
	put(before, 'B1', { value: 2, formula: 'A1*2' });
	put(after, 'B1', { value: cellError('#VALUE!'), formula: 'A1*3' });
	put(before, 'C1', { value: 3, formula: 'A1+2' });
	put(after, 'C1', { value: 4, formula: 'A1+2' });
	const percent = internStyle(after, patchStyle(defaultCellStyle(), { numFmt: '0%' }));
	put(before, 'A2', { value: 0.5 });
	put(after, 'A2', { value: 0.5, styleId: percent });
	for (const ref of ['A4', 'B4', 'A5', 'B5']) put(after, ref, { value: true });
	return [before, after];
}

describe('comparisonText', () => {
	it('lists workbook, structure and cell differences in order, grouping neighbouring cells', () => {
		const [before, after] = pair();
		expect(comparisonText(compareWorkbooks(before, after))).toBe(
			[
				'Workbook',
				'  Sheet "Old" removed from position 2',
				'  Sheet "Extra" added at position 3',
				'  Sheet "Data" renamed to "Inputs"',
				'  Sheet "Notes" visible -> hidden',
				'  Name Local ("Notes") added: =1 (hidden)',
				'  Name Rate changed: =Data!$A$1 -> =Inputs!$A$2',
				'Sheet "Inputs" (was "Data")',
				'  Merge added D1:E1',
				'  A1 entered value: 1 -> "one"',
				'  B1 formula: =A1*2 -> =A1*3',
				'  C1 calculated value: 3 -> 4 (=A1+2)',
				'  A2 format (number format): "General" -> "0%"',
				'  A4:B5 entered value: 4 cells',
				'Not compared: rowAndColumnLayout, rowAndColumnFormats, comments, hyperlinks, drawings, pivotTables, conditionalFormatExtensions, sparklines, slicers, autoFilter, sheetViews, pageSetup, sheetProtection, documentProperties, theme, workbookSettings, macrosAndUnmodelledParts.',
				'',
			].join('\n'),
		);
	});

	it('prints one line per cell on request', () => {
		const [before, after] = pair();
		const text = comparisonText(compareWorkbooks(before, after, { categories: ['enteredValue'] }), {
			cells: true,
		});
		expect(text.split('\n').slice(0, 7)).toEqual([
			'Sheet "Inputs" (was "Data")',
			'  A1 entered value: 1 -> "one"',
			'  A4 entered value: (empty) -> TRUE',
			'  B4 entered value: (empty) -> TRUE',
			'  A5 entered value: (empty) -> TRUE',
			'  B5 entered value: (empty) -> TRUE',
			'Not compared: formula, calculatedValue, format, sheet, name, merge, table, dataValidation, conditionalFormat, rowAndColumnLayout, rowAndColumnFormats, comments, hyperlinks, drawings, pivotTables, conditionalFormatExtensions, sparklines, slicers, autoFilter, sheetViews, pageSetup, sheetProtection, documentProperties, theme, workbookSettings, macrosAndUnmodelledParts.',
		]);
	});

	it('keeps every difference on one line and shows what a rule or name changed', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		put(before, 'A1', { value: 1, formula: 'IF(B1,\n1,2)' });
		put(after, 'A1', { value: 1, formula: 'IF(B1,\r\n2,2)' });
		before.definedNames = [{ name: 'Total', formula: '1' }];
		after.definedNames = [{ name: 'TOTAL', formula: '1' }];
		before.sheets[0]!.dataValidations = [
			{ ranges: [R('C2')], type: 'whole', operator: 'between', formula1: '1', formula2: '9' },
		];
		after.sheets[0]!.dataValidations = [
			{ ranges: [R('C2')], type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
		];
		const lines = comparisonText(compareWorkbooks(before, after)).split('\n');
		expect(lines.slice(0, 6)).toEqual([
			'Workbook',
			'  Name Total renamed to TOTAL',
			'Sheet "Sheet1"',
			'  Data validation removed C2 (type whole, operator between, formula1 "1", formula2 "9")',
			'  Data validation added C2 (type whole, operator between, formula1 "1", formula2 "10")',
			'  A1 formula: =IF(B1,\\n1,2) -> =IF(B1,\\r\\n2,2)',
		]);
	});

	it('escapes line and paragraph separators in values and formulas', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		put(before, 'A1', { value: 'one' });
		put(after, 'A1', { value: 'one\u2028Sheet "Fake"\u0085two' });
		put(before, 'B1', { value: 1, formula: '1' });
		put(after, 'B1', { value: 2, formula: '1+1\u2029' });
		const text = comparisonText(compareWorkbooks(before, after));
		expect(text).not.toMatch(/[\u0085\u2028\u2029]/u);
		expect(text).toContain('A1 entered value: "one" -> "one\\u2028Sheet \\"Fake\\"\\u0085two"');
		expect(text).toContain('B1 formula: =1 -> =1+1\\u2029');
	});

	it('shows a setting turned off, not only the ones turned on', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		const list = { ranges: [R('A1')], type: 'list' as const, formula1: '"a,b"' };
		before.sheets[0]!.dataValidations = [list];
		after.sheets[0]!.dataValidations = [{ ...list, showDropDown: false }];
		const bar = (showValue: boolean) => ({
			ranges: [R('B1:B9')],
			rules: [
				{
					type: 'dataBar' as const,
					min: { type: 'min' as const },
					max: { type: 'max' as const },
					color: { rgb: 'FF638EC6' },
					priority: 1,
					showValue,
				},
			],
		});
		before.sheets[0]!.conditionalFormats = [bar(true)];
		after.sheets[0]!.conditionalFormats = [bar(false)];
		const lines = comparisonText(compareWorkbooks(before, after)).split('\n').slice(1, 5);
		expect(lines).toEqual([
			'  Data validation removed A1 (type list, formula1 "\\"a,b\\"")',
			'  Data validation added A1 (type list, formula1 "\\"a,b\\"", showDropDown false)',
			expect.stringMatching(
				/^ {2}Conditional format rule removed B1:B9 \(type dataBar, color .*\)$/,
			),
			expect.stringMatching(/^ {2}Conditional format rule added B1:B9 \(.*showValue false.*\)$/),
		]);
		expect([lines[2]?.endsWith(', showValue)'), lines[3]?.endsWith(', showValue false)')]).toEqual([
			true,
			true,
		]);
	});

	it('says when there is nothing to report and when the report stopped early', () => {
		const book = createWorkbook();
		expect(comparisonText(compareWorkbooks(book, structuredClone(book)))).toMatch(
			/^No differences\.\nNot compared: rowAndColumnLayout, /,
		);
		const [before, after] = pair();
		const lines = comparisonText(compareWorkbooks(before, after, { maxDifferences: 2 })).split(
			'\n',
		);
		expect(lines.slice(0, 4)).toEqual([
			'Workbook',
			'  Sheet "Old" removed from position 2',
			'  Sheet "Extra" added at position 3',
			'Stopped after 2 differences; there may be more.',
		]);
	});
});
