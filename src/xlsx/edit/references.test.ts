import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const cell = (wb: Workbook, row: number, col: number, sheet = 0) =>
	getCell(wb.sheets[sheet]!, row, col);
const range = (r1: number, c1: number, r2 = r1, c2 = c1) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

function twoSheets() {
	const wb = createWorkbook({ sheets: ['Data', 'Report'] });
	const s = createEditSession(wb);
	return { wb, s };
}

describe('cut and paste', () => {
	it('rewrites references to the moved cells on every sheet, undoably', () => {
		const { wb, s } = twoSheets();
		s.setCellValue(0, 0, 0, 10);
		s.setCellInput(0, 0, 1, '=A1*2');
		s.setCellInput(1, 0, 0, '=Data!A1+1');
		s.setDefinedName({ name: 'Base', formula: 'Data!$A$1' });
		s.paste(0, { row: 4, col: 2 }, s.cut(0, range(0, 0)));
		expect(cell(wb, 0, 1)?.formula).toBe('C5*2');
		expect(cell(wb, 0, 0, 1)?.formula).toBe('Data!C5+1');
		expect(wb.definedNames[0]?.formula).toBe('Data!$C$5');
		expect(cell(wb, 0, 0, 1)?.value).toBe(11);
		s.undo();
		expect(cell(wb, 0, 1)?.formula).toBe('A1*2');
		expect(cell(wb, 0, 0, 1)?.formula).toBe('Data!A1+1');
		expect(wb.definedNames[0]?.formula).toBe('Data!$A$1');
	});
	it('moves a formula block to another sheet keeping its meaning', () => {
		const { wb, s } = twoSheets();
		s.setCellValue(0, 0, 0, 3);
		s.setCellInput(0, 0, 1, '=A1+A2');
		s.paste(1, { row: 0, col: 0 }, s.cut(0, range(0, 1)));
		expect(cell(wb, 0, 0, 1)?.formula).toBe('Data!A1+Data!A2');
		expect(cell(wb, 0, 0, 1)?.value).toBe(3);
	});
});

describe('sheet operations rewrite references', () => {
	it('deleteSheet turns references to it into #REF! everywhere', () => {
		const { wb, s } = twoSheets();
		s.setCellValue(0, 1, 0, 1);
		s.createTable(0, range(0, 0, 2, 1), true);
		s.setCellInput(1, 0, 0, '=Data!A2');
		s.setCellInput(1, 1, 0, '=SUM(Table1[Column2])');
		s.setDefinedName({ name: 'Rate', formula: 'Data!$B$2' });
		s.setDataValidation(1, { ranges: [], type: 'list', formula1: 'Data!$A$1:$A$3' }, range(5, 5));
		s.deleteSheet(0);
		const report = wb.sheets[0]!;
		expect(getCell(report, 0, 0)?.formula).toBe('#REF!');
		expect(getCell(report, 0, 0)?.value).toEqual({ error: '#REF!' });
		expect(getCell(report, 1, 0)?.formula).toBe('SUM(#REF!)');
		expect(wb.definedNames.find((n) => n.name === 'Rate')?.formula).toBe('#REF!');
		expect(report.dataValidations[0]?.formula1).toBe('#REF!');
		s.undo();
		expect(cell(wb, 0, 0, 1)?.formula).toBe('Data!A2');
		expect(cell(wb, 0, 0, 1)?.value).toBe(1);
	});
	it('renameSheet rewrites formulas, names, validation, CF, charts and hyperlinks', () => {
		const { wb, s } = twoSheets();
		s.setCellInput(1, 0, 0, '=Data!A1');
		s.setDefinedName({ name: 'Rate', formula: 'Data!$B$2' });
		s.setHyperlink(1, range(2, 2), { location: 'Data!A1' });
		s.addConditionalFormat(1, {
			ranges: [range(0, 0)],
			rules: [{ type: 'expression', formula: 'Data!$A$1>0', style: {}, priority: 1 }],
		});
		wb.sheets[1]!.drawings.push({
			kind: 'chart',
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
			chartType: 'bar',
			series: [{ valuesRef: 'Data!$A$1:$A$3', categories: [], values: [] }],
			showLegend: true,
		});
		s.renameSheet(0, 'My Data');
		expect(cell(wb, 0, 0, 1)?.formula).toBe("'My Data'!A1");
		expect(wb.definedNames[0]?.formula).toBe("'My Data'!$B$2");
		expect(wb.sheets[1]!.hyperlinks[0]?.location).toBe("'My Data'!A1");
		const rule = wb.sheets[1]!.conditionalFormats[0]?.rules[0];
		expect(rule?.type === 'expression' && rule.formula).toBe("'My Data'!$A$1>0");
		const chart = wb.sheets[1]!.drawings[0];
		expect(chart?.kind === 'chart' && chart.series[0]?.valuesRef).toBe("'My Data'!$A$1:$A$3");
	});
	it('duplicateSheet gives tables unique names and rewrites the copy only', async () => {
		const { wb, s } = twoSheets();
		s.setRangeValues(0, { row: 0, col: 0 }, [
			['Item', 'Qty'],
			['a', 1],
			['b', 2],
		]);
		const table = s.createTable(0, range(0, 0, 2, 1), true);
		s.setCellInput(0, 0, 3, `=SUM(${table.name}[Qty])`);
		const copy = s.duplicateSheet(0);
		const copied = wb.sheets[copy]!;
		const name = copied.tables[0]!.name;
		expect(name).not.toBe(table.name);
		expect(getCell(copied, 0, 3)?.formula).toBe(`SUM(${name}[Qty])`);
		expect(cell(wb, 0, 3)?.formula).toBe(`SUM(${table.name}[Qty])`);
		expect(getCell(copied, 0, 3)?.value).toBe(3);
		const again = s.duplicateSheet(copy);
		expect(new Set(wb.sheets.flatMap((sh) => sh.tables.map((t) => t.name))).size).toBe(3);
		expect(wb.sheets[again]!.tables[0]!.name).toMatch(/^Table1_\d$/);
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets.flatMap((sh) => sh.tables.map((t) => t.name))).toEqual(
			wb.sheets.flatMap((sh) => sh.tables.map((t) => t.name)),
		);
	});
	it('new sheets use the session base name', () => {
		const wb = createWorkbook({ sheets: ['Feuille1'] });
		const s = createEditSession(wb, { defaultSheetBase: 'Feuille' });
		expect(wb.sheets[s.addSheet()]?.name).toBe('Feuille2');
	});
});

describe('sort', () => {
	it('moves comments and single-row hyperlinks with their rows', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb);
		s.setRangeValues(0, { row: 0, col: 0 }, [[3], [1], [2]]);
		s.setComment(0, { row: 0, col: 0 }, 'three', 'me');
		s.setHyperlink(0, range(1, 0), { target: 'https://example.com/one' });
		s.setRowHeight(0, [0], 30);
		s.sort(0, range(0, 0, 2, 0), [{ col: 0 }], false);
		expect(wb.sheets[0]!.comments[0]?.address).toEqual({ row: 2, col: 0 });
		expect(wb.sheets[0]!.hyperlinks[0]?.range).toEqual(range(0, 0));
		expect(wb.sheets[0]!.rowInfo.get(0)?.height).toBe(30);
		s.undo();
		expect(wb.sheets[0]!.comments[0]?.address).toEqual({ row: 0, col: 0 });
	});
});

describe('structural edits recalculate', () => {
	it('after row inserts, name edits and table creation', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb);
		s.setCellValue(0, 0, 0, 2);
		s.setCellInput(0, 0, 1, '=A1*Rate');
		s.setDefinedName({ name: 'Rate', formula: '10' });
		expect(cell(wb, 0, 1)?.value).toBe(20);
		s.setDefinedName({ name: 'Rate', formula: '3' });
		expect(cell(wb, 0, 1)?.value).toBe(6);
		s.deleteDefinedName('Rate');
		expect(cell(wb, 0, 1)?.value).toEqual({ error: '#NAME?' });
		s.undo();
		expect(cell(wb, 0, 1)?.value).toBe(6);
		s.insertRows(0, 0, 1);
		expect(cell(wb, 1, 1)?.formula).toBe('A2*Rate');
		expect(cell(wb, 1, 1)?.value).toBe(6);
		putCell(wb.sheets[0]!, 5, 0, { value: 'Total' });
		s.setCellInput(0, 6, 0, '=ROWS(Table1)');
		s.createTable(0, range(1, 0, 3, 1), true);
		expect(cell(wb, 6, 0)?.value).toBe(2);
	});
});
