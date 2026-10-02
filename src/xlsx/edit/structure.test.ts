import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { ConditionalFormat, Workbook, Worksheet } from '../model.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const A = (ref: string) => {
	const a = parseAddress(ref);
	if (!a) throw new Error(ref);
	return a;
};
const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const ws = (wb: Workbook, i = 0): Worksheet => {
	const sheet = wb.sheets[i];
	if (!sheet) throw new Error('sheet');
	return sheet;
};
const cell = (wb: Workbook, ref: string, i = 0) => getCell(ws(wb, i), A(ref).row, A(ref).col);
const setup = (sheets = ['Sheet1', 'Sheet2']) => {
	const wb = createWorkbook({ sheets });
	return { wb, s: createEditSession(wb, { recalc: false }) };
};
const cf = (ref: string): ConditionalFormat => ({
	ranges: [R(ref)],
	rules: [{ type: 'expression', formula: 'A1>0', style: { font: { bold: true } }, priority: 1 }],
});

describe('insertRows / deleteRows', () => {
	it('moves cells down and back up', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [[1], [2], [3]]);
		s.insertRows(0, 1, 2);
		expect(cell(wb, 'A1')?.value).toBe(1);
		expect(cell(wb, 'A2')).toBeUndefined();
		expect(cell(wb, 'A4')?.value).toBe(2);
		expect(cell(wb, 'A5')?.value).toBe(3);
		s.deleteRows(0, 1, 2);
		expect(cell(wb, 'A2')?.value).toBe(2);
		expect(cell(wb, 'A3')?.value).toBe(3);
	});
	it('deletes the cells in the deleted rows', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [[1], [2], [3]]);
		s.deleteRows(0, 0, 2);
		expect(cell(wb, 'A1')?.value).toBe(3);
		expect(ws(wb).rows.size).toBe(1);
	});
	it('shifts formulas on the same sheet', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 1, '=A5*2');
		s.insertRows(0, 2, 3);
		expect(cell(wb, 'B1')?.formula).toBe('A8*2');
	});
	it('shifts formulas on other sheets that point at the edited sheet', () => {
		const { wb, s } = setup();
		s.setCellInput(1, 0, 0, '=Sheet1!A5+A5');
		s.insertRows(0, 0, 1);
		expect(cell(wb, 'A1', 1)?.formula).toBe('Sheet1!A6+A5');
	});
	it('turns references to deleted rows into #REF!', () => {
		const { wb, s } = setup();
		s.setCellInput(1, 0, 0, '=Sheet1!A2');
		s.deleteRows(0, 1, 1);
		expect(cell(wb, 'A1', 1)?.formula).toContain('#REF!');
	});
	it('widens ranges that the insert falls inside', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 3, '=SUM(A1:A10)');
		s.insertRows(0, 4, 2);
		expect(cell(wb, 'D1')?.formula).toBe('SUM(A1:A12)');
		s.deleteRows(0, 2, 3);
		expect(cell(wb, 'D1')?.formula).toBe('SUM(A1:A9)');
	});
	it('undoes an insert across every sheet', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 3, 0, 'x');
		s.setCellInput(1, 0, 0, '=Sheet1!A4');
		s.insertRows(0, 0, 1);
		s.undo();
		expect(cell(wb, 'A4')?.value).toBe('x');
		expect(cell(wb, 'A1', 1)?.formula).toBe('Sheet1!A4');
	});
	it('moves merges and drops merges in deleted rows', () => {
		const { wb, s } = setup();
		s.merge(0, R('A3:B4'), 'merge');
		s.merge(0, R('D1:E1'), 'merge');
		s.insertRows(0, 0, 1);
		expect(ws(wb).merges).toEqual([R('A4:B5'), R('D2:E2')]);
		s.deleteRows(0, 1, 1);
		expect(ws(wb).merges).toEqual([R('A3:B4')]);
	});
	it('shrinks a merge to nothing when it collapses to one cell', () => {
		const { wb, s } = setup();
		s.merge(0, R('A1:A2'), 'merge');
		s.deleteRows(0, 1, 1);
		expect(ws(wb).merges).toEqual([]);
	});
	it('moves conditional-format and validation ranges', () => {
		const { wb, s } = setup();
		s.addConditionalFormat(0, cf('B2:B5'));
		s.setDataValidation(
			0,
			{ ranges: [], type: 'whole', operator: 'greaterThan', formula1: '0' },
			R('C3:C4'),
		);
		s.insertRows(0, 0, 2);
		expect(ws(wb).conditionalFormats[0]?.ranges).toEqual([R('B4:B7')]);
		expect(ws(wb).dataValidations[0]?.ranges).toEqual([R('C5:C6')]);
		s.deleteRows(0, 3, 4);
		expect(ws(wb).conditionalFormats).toHaveLength(0);
		expect(ws(wb).dataValidations).toHaveLength(0);
	});
	it('moves hyperlinks and comments', () => {
		const { wb, s } = setup();
		s.setHyperlink(0, R('A2'), { target: 'https://example.com' });
		s.setComment(0, A('B3'), 'hi', 'me');
		s.insertRows(0, 1, 1);
		expect(ws(wb).hyperlinks[0]?.range).toEqual(R('A3'));
		expect(ws(wb).comments[0]?.address).toEqual(A('B4'));
		s.deleteRows(0, 3, 1);
		expect(ws(wb).comments).toHaveLength(0);
	});
	it('moves and grows tables', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A2'), [
			['h1', 'h2'],
			[1, 2],
			[3, 4],
		]);
		s.createTable(0, R('A2:B4'), true);
		s.insertRows(0, 0, 1);
		expect(ws(wb).tables[0]?.range).toEqual(R('A3:B5'));
		s.insertRows(0, 4, 2);
		expect(ws(wb).tables[0]?.range).toEqual(R('A3:B7'));
		s.deleteRows(0, 0, 10);
		expect(ws(wb).tables).toHaveLength(0);
	});
	it('moves drawing anchors', () => {
		const { wb, s } = setup();
		ws(wb).drawings.push({
			kind: 'unsupported',
			description: 'shape',
			anchor: {
				from: { row: 2, col: 1, rowOffset: 10, colOffset: 0 },
				to: { row: 5, col: 3, rowOffset: 0, colOffset: 0 },
			},
		});
		s.insertRows(0, 0, 2);
		expect(ws(wb).drawings[0]?.anchor.from.row).toBe(4);
		expect(ws(wb).drawings[0]?.anchor.to?.row).toBe(7);
		s.deleteRows(0, 3, 2);
		expect(ws(wb).drawings[0]?.anchor.from).toMatchObject({ row: 3, rowOffset: 0 });
		expect(ws(wb).drawings[0]?.anchor.to?.row).toBe(5);
	});
	it('moves row formats and copies the format of the row above', () => {
		const { wb, s } = setup();
		ws(wb).rowInfo.set(3, { height: 40, customHeight: true });
		s.setCellValue(0, 0, 0, 'h');
		s.applyStyle(0, [R('A1')], { font: { bold: true } });
		s.insertRows(0, 1, 1);
		expect(ws(wb).rowInfo.get(4)?.height).toBe(40);
		expect(ws(wb).rowInfo.has(3)).toBe(false);
		expect(styleAt(wb, cell(wb, 'A2')?.styleId).font.bold).toBe(true);
		expect(cell(wb, 'A2')?.value).toBeNull();
	});
	it('updates defined names, the auto-filter, the freeze pane and the print area', () => {
		const { wb, s } = setup();
		s.setDefinedName({ name: 'Data', formula: 'Sheet1!$A$1:$A$10' });
		s.setAutoFilter(0, R('A1:C10'));
		s.setFreeze(0, { rows: 2, cols: 0 });
		ws(wb).pageSetup = { printArea: R('A1:C10') };
		s.insertRows(0, 1, 1);
		expect(wb.definedNames[0]?.formula).toBe('Sheet1!$A$1:$A$11');
		expect(ws(wb).autoFilter?.range).toEqual(R('A1:C11'));
		expect(ws(wb).view.freeze).toEqual({ rows: 3, cols: 0 });
		expect(ws(wb).pageSetup?.printArea).toEqual(R('A1:C11'));
		s.deleteRows(0, 0, 2);
		expect(ws(wb).view.freeze).toEqual({ rows: 1, cols: 0 });
	});
	it('validates the span', () => {
		const { s } = setup();
		expect(() => s.insertRows(0, -1, 1)).toThrow(RangeError);
		expect(() => s.deleteRows(0, 0, 0)).toThrow(RangeError);
	});
});

describe('insertColumns / deleteColumns', () => {
	it('moves cells right and back', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [[1, 2, 3]]);
		s.insertColumns(0, 1, 1);
		expect(cell(wb, 'C1')?.value).toBe(2);
		expect(cell(wb, 'B1')).toBeUndefined();
		s.deleteColumns(0, 0, 2);
		expect(cell(wb, 'A1')?.value).toBe(2);
		expect(cell(wb, 'B1')?.value).toBe(3);
	});
	it('shifts formulas on every sheet', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 5, 0, '=SUM(B1:D1)');
		s.setCellInput(1, 0, 0, "='Sheet1'!C1");
		s.insertColumns(0, 2, 1);
		expect(cell(wb, 'A6')?.formula).toBe('SUM(B1:E1)');
		expect(cell(wb, 'A1', 1)?.formula?.replace(/'/g, '')).toBe('Sheet1!D1');
	});
	it('updates column formats', () => {
		const { wb, s } = setup();
		s.setColumnWidth(0, [2], 25);
		s.insertColumns(0, 0, 2);
		expect(ws(wb).columns).toEqual([{ min: 4, max: 4, width: 25, customWidth: true }]);
		s.deleteColumns(0, 4, 1);
		expect(ws(wb).columns).toEqual([]);
	});
	it('adds and removes table columns', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			['a', 'b', 'c'],
			[1, 2, 3],
		]);
		const table = s.createTable(0, R('A1:C2'), true);
		expect(table.columns.map((c) => c.name)).toEqual(['a', 'b', 'c']);
		s.insertColumns(0, 1, 1);
		const t = ws(wb).tables[0];
		expect(t?.range).toEqual(R('A1:D2'));
		expect(t?.columns.map((c) => c.name)).toEqual(['a', 'Column1', 'b', 'c']);
		expect(cell(wb, 'B1')?.value).toBe('Column1');
		s.deleteColumns(0, 2, 2);
		expect(ws(wb).tables[0]?.columns.map((c) => c.name)).toEqual(['a', 'Column1']);
	});
	it('moves auto-filter column filters', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			['h1', 'h2', 'h3'],
			['x', 'y', 'z'],
		]);
		s.setAutoFilter(0, R('A1:C2'));
		s.filterColumn(0, 2, ['z']);
		s.deleteColumns(0, 0, 1);
		expect(ws(wb).autoFilter?.range).toEqual(R('A1:B2'));
		expect(ws(wb).autoFilter?.columns).toEqual([{ offset: 1, values: ['z'] }]);
	});
	it('updates the frozen columns', () => {
		const { wb, s } = setup();
		s.setFreeze(0, { rows: 0, cols: 2 });
		s.insertColumns(0, 0, 1);
		expect(ws(wb).view.freeze).toEqual({ rows: 0, cols: 3 });
		s.insertColumns(0, 5, 1);
		expect(ws(wb).view.freeze).toEqual({ rows: 0, cols: 3 });
	});
});

describe('insert / delete cells with shift', () => {
	it('shifts cells down within the columns of the range', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			[1, 2],
			[3, 4],
		]);
		s.insertCellsShift(0, R('A1'), 'down');
		expect(cell(wb, 'A1')).toBeUndefined();
		expect(cell(wb, 'A2')?.value).toBe(1);
		expect(cell(wb, 'A3')?.value).toBe(3);
		expect(cell(wb, 'B1')?.value).toBe(2);
		s.undo();
		expect(cell(wb, 'A1')?.value).toBe(1);
	});
	it('shifts cells right within the rows of the range', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			[1, 2],
			[3, 4],
		]);
		s.insertCellsShift(0, R('A2'), 'right');
		expect(cell(wb, 'B2')?.value).toBe(3);
		expect(cell(wb, 'C2')?.value).toBe(4);
		expect(cell(wb, 'A1')?.value).toBe(1);
	});
	it('deletes cells and pulls the rest up or left', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			[1, 2],
			[3, 4],
			[5, 6],
		]);
		s.deleteCellsShift(0, R('A1:A2'), 'up');
		expect(cell(wb, 'A1')?.value).toBe(5);
		expect(cell(wb, 'B1')?.value).toBe(2);
		s.deleteCellsShift(0, R('A1'), 'left');
		expect(cell(wb, 'A1')?.value).toBe(2);
	});
	it('treats whole rows as a row insert', () => {
		const { wb, s } = setup();
		s.setCellInput(1, 0, 0, '=Sheet1!A1');
		s.insertCellsShift(0, R('1:1'), 'down');
		expect(cell(wb, 'A1', 1)?.formula).toBe('Sheet1!A2');
	});
	it('moves references that lie inside the shifted band', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [[1], [2]]);
		s.setCellInput(0, 0, 3, '=A2+SUM(A1:A2)+B2');
		s.setCellInput(1, 0, 0, '=Sheet1!$A$2');
		s.insertCellsShift(0, R('A1'), 'down');
		expect(cell(wb, 'D1')?.formula).toBe('A3+SUM(A2:A3)+B2');
		expect(cell(wb, 'A1', 1)?.formula).toBe('Sheet1!$A$3');
		s.undo();
		expect(cell(wb, 'A1', 1)?.formula).toBe('Sheet1!$A$2');
	});
	it('leaves references straddling the band and marks deleted ones #REF!', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 3, '=SUM(A2:B2)+A1');
		s.deleteCellsShift(0, R('A1'), 'up');
		expect(cell(wb, 'D1')?.formula).toBe('SUM(A2:B2)+#REF!');
	});
	it('refuses to split a merged range', () => {
		const { s } = setup();
		s.merge(0, R('A3:B3'), 'merge');
		expect(() => s.insertCellsShift(0, R('A1'), 'down')).toThrow(/merged/);
	});
});
