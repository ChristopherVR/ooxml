import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell, putCell } from '../cells';
import { loadXlsx } from '../read';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write';
import { createEditSession } from './session';

const range = (text: string) => parseRange(text)!;

function setup() {
	const workbook = createWorkbook({ sheets: ['Data', 'Report'] });
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Item', 'Amount', 'Double'],
		['first', 10, null],
		['second', 20, null],
	]);
	session.createTable(0, range('A1:C3'), true);
	session.setCellInput(0, 1, 2, '=[@Amount]*2');
	session.setCellInput(1, 0, 0, '=SUM(Table1[Amount])');
	session.setDefinedName({ name: 'TotalAmount', formula: 'SUM(Table1[Amount])' });
	workbook.sheets[0]!.tables[0]!.columns[2]!.calculatedColumnFormula = '[@Amount]*2';
	return { workbook, session };
}

describe('table header edits', () => {
	it('finds references added between successive header edits', () => {
		const { workbook, session } = setup();
		session.setCellInput(0, 0, 1, 'Cost');
		putCell(workbook.sheets[1]!, 5, 0, { value: null, formula: 'SUM(Table1[Cost])' });
		session.setCellInput(0, 0, 1, 'Price');
		expect(getCell(workbook.sheets[1]!, 5, 0)).toMatchObject({
			formula: 'SUM(Table1[Price])',
			value: 30,
		});
		session.undo();
		expect(getCell(workbook.sheets[1]!, 5, 0)?.formula).toBe('SUM(Table1[Cost])');
	});

	it('keeps duplicate names unique and rewrites references to the resulting name', () => {
		const { workbook, session } = setup();
		session.setCellInput(0, 0, 1, 'Item');
		expect(workbook.sheets[0]!.tables[0]!.columns[1]!.name).toBe('Item2');
		expect(getCell(workbook.sheets[0]!, 0, 1)?.value).toBe('Item2');
		expect(getCell(workbook.sheets[1]!, 0, 0)).toMatchObject({
			formula: 'SUM(Table1[Item2])',
			value: 30,
		});
	});

	it('rewrites validation and conditional formulas while leaving other tables and strings alone', () => {
		const { workbook, session } = setup();
		session.setRangeValues(1, { row: 2, col: 0 }, [['Amount'], [7]]);
		session.createTable(1, range('A3:A4'), true);
		session.setCellInput(1, 2, 2, '=SUM(Table2[Amount])');
		session.setCellInput(1, 3, 2, '="Table1[Amount]"');
		const sheet = workbook.sheets[1]!;
		sheet.dataValidations.push({
			ranges: [range('D1')],
			type: 'custom',
			formula1: 'SUM(Table1[Amount])>0',
		});
		sheet.conditionalFormats.push({
			ranges: [range('D1')],
			rules: [{ type: 'expression', priority: 1, style: {}, formula: 'SUM(Table1[Amount])>0' }],
		});
		session.setCellInput(0, 0, 1, 'Cost');
		expect(sheet.dataValidations[0]!.formula1).toBe('SUM(Table1[Cost])>0');
		expect(sheet.conditionalFormats[0]!.rules[0]).toMatchObject({ formula: 'SUM(Table1[Cost])>0' });
		expect(getCell(sheet, 2, 2)).toMatchObject({ formula: 'SUM(Table2[Amount])', value: 7 });
		expect(getCell(sheet, 3, 2)?.formula).toBe('"Table1[Amount]"');
		session.undo();
		expect(sheet.dataValidations[0]!.formula1).toBe('SUM(Table1[Amount])>0');
	});

	it('renames references across the workbook and restores them with undo and redo', () => {
		const { workbook, session } = setup();
		expect(getCell(workbook.sheets[1]!, 0, 0)?.value).toBe(30);
		session.setCellInput(0, 0, 1, 'Cost');
		const assertNames = (name: string) => {
			expect(workbook.sheets[0]!.tables[0]!.columns[1]!.name).toBe(name);
			expect(getCell(workbook.sheets[1]!, 0, 0)).toMatchObject({
				formula: `SUM(Table1[${name}])`,
				value: 30,
			});
			expect(getCell(workbook.sheets[0]!, 1, 2)).toMatchObject({
				formula: `[@${name}]*2`,
				value: 20,
			});
			expect(workbook.definedNames[0]!.formula).toBe(`SUM(Table1[${name}])`);
			expect(workbook.sheets[0]!.tables[0]!.columns[2]!.calculatedColumnFormula).toBe(
				`[@${name}]*2`,
			);
		};
		assertNames('Cost');
		session.undo();
		assertNames('Amount');
		session.redo();
		assertNames('Cost');
	});

	it.each(['value', 'range', 'replace', 'paste', 'fill'] as const)(
		'updates references when a header changes through %s',
		(method) => {
			const { workbook, session } = setup();
			switch (method) {
				case 'value':
					session.setCellValue(0, 0, 1, 'Cost');
					break;
				case 'range':
					session.setRangeValues(0, { row: 0, col: 1 }, [['Cost']]);
					break;
				case 'replace':
					session.replaceAll(
						{ sheet: 0, range: range('B1'), text: 'Amount', wholeCell: true },
						'Cost',
					);
					break;
				case 'paste':
					session.paste(0, { row: 0, col: 1 }, 'Cost');
					break;
				case 'fill':
					session.setCellValue(0, 0, 3, 'Cost');
					session.fill(0, range('D1'), range('B1'), 'copy');
					break;
			}
			expect(getCell(workbook.sheets[1]!, 0, 0)).toMatchObject({
				formula: 'SUM(Table1[Cost])',
				value: 30,
			});
			session.undo();
			expect(getCell(workbook.sheets[1]!, 0, 0)).toMatchObject({
				formula: 'SUM(Table1[Amount])',
				value: 30,
			});
			session.redo();
			expect(getCell(workbook.sheets[1]!, 0, 0)).toMatchObject({
				formula: 'SUM(Table1[Cost])',
				value: 30,
			});
		},
	);

	it('saves the renamed header and formulas together', async () => {
		const { workbook, session } = setup();
		session.setCellInput(0, 0, 1, 'Cost');
		const loaded = await loadXlsx(await saveXlsx(workbook));
		createEditSession(loaded, { autoRowHeight: false });
		expect(loaded.sheets[0]!.tables[0]!.columns[1]!.name).toBe('Cost');
		expect(getCell(loaded.sheets[1]!, 0, 0)).toMatchObject({
			formula: 'SUM(Table1[Cost])',
			value: 30,
		});
	});
});
