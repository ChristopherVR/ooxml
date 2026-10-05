import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { ConditionalFormat, Workbook, Worksheet } from '../model.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';
import { uniqueHeaders } from './tables.js';
import { validateDefinedName } from './view.js';

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
const cell = (wb: Workbook, ref: string) => getCell(ws(wb), A(ref).row, A(ref).col);
const setup = () => {
	const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
	return { wb, s: createEditSession(wb, { recalc: false }) };
};
const bold = (ref: string, priority = 1): ConditionalFormat => ({
	ranges: [R(ref)],
	rules: [
		{
			type: 'cellIs',
			operator: 'greaterThan',
			formulas: ['0'],
			style: { font: { bold: true } },
			priority,
		},
	],
});

describe('comments', () => {
	it('adds, edits and deletes a comment', () => {
		const { wb, s } = setup();
		s.setComment(0, A('B2'), 'first', 'Ann');
		expect(ws(wb).comments).toEqual([{ address: A('B2'), author: 'Ann', text: 'first' }]);
		s.setComment(0, A('B2'), 'second', 'Bob');
		expect(ws(wb).comments).toEqual([{ address: A('B2'), author: 'Bob', text: 'second' }]);
		s.setComment(0, A('B2'), undefined, 'Bob');
		expect(ws(wb).comments).toEqual([]);
		s.undo();
		expect(ws(wb).comments[0]?.text).toBe('second');
		expect(s.undoLabel()).toBe('Edit comment');
	});
	it('ignores deleting a missing comment', () => {
		const { s } = setup();
		s.setComment(0, A('A1'), undefined, 'x');
		expect(s.canUndo()).toBe(false);
	});
});

describe('hyperlinks', () => {
	it('adds a link, fills an empty cell and styles it', () => {
		const { wb, s } = setup();
		s.setHyperlink(0, R('A1'), { target: 'https://example.com', display: 'Example' });
		expect(ws(wb).hyperlinks).toEqual([
			{ target: 'https://example.com', display: 'Example', range: R('A1') },
		]);
		expect(cell(wb, 'A1')?.value).toBe('Example');
		const font = styleAt(wb, cell(wb, 'A1')?.styleId).font;
		expect(font.underline).toBe('single');
		expect(font.color).toEqual({ theme: 10 });
	});
	it('keeps existing cell text and replaces overlapping links', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 'mine');
		s.setHyperlink(0, R('A1'), { target: 'https://a.example' });
		s.setHyperlink(0, R('A1:B1'), { location: 'Sheet2!A1' });
		expect(cell(wb, 'A1')?.value).toBe('mine');
		expect(ws(wb).hyperlinks).toEqual([{ location: 'Sheet2!A1', range: R('A1:B1') }]);
	});
	it('removes links and rejects empty ones', () => {
		const { wb, s } = setup();
		s.setHyperlink(0, R('A1'), { target: 'https://a.example' });
		s.setHyperlink(0, R('A1'), undefined);
		expect(ws(wb).hyperlinks).toEqual([]);
		expect(() => s.setHyperlink(0, R('A1'), {})).toThrow();
	});
});

describe('conditional formats', () => {
	it('gives new rules the top priority', () => {
		const { wb, s } = setup();
		s.addConditionalFormat(0, bold('A1:A10'));
		s.addConditionalFormat(0, bold('B1:B10', 7));
		const [first, second] = ws(wb).conditionalFormats;
		expect(first?.rules[0]?.priority).toBe(2);
		expect(second?.rules[0]?.priority).toBe(1);
	});
	it('clears rules from part of their range', () => {
		const { wb, s } = setup();
		s.addConditionalFormat(0, bold('A1:A10'));
		s.clearConditionalFormats(0, R('A1:A5'));
		expect(ws(wb).conditionalFormats[0]?.ranges).toEqual([R('A6:A10')]);
		s.clearConditionalFormats(0, R('A:A'));
		expect(ws(wb).conditionalFormats).toEqual([]);
		s.undo();
		expect(ws(wb).conditionalFormats).toHaveLength(1);
	});
	it('clears every rule of the sheet', () => {
		const { wb, s } = setup();
		s.addConditionalFormat(0, bold('A1'));
		s.addConditionalFormat(0, bold('C1'));
		s.clearConditionalFormats(0);
		expect(ws(wb).conditionalFormats).toEqual([]);
	});
});

describe('data validation', () => {
	it('replaces the part of existing rules the new range covers', () => {
		const { wb, s } = setup();
		s.setDataValidation(0, { ranges: [], type: 'list', formula1: '"a,b"' }, R('A1:A10'));
		s.setDataValidation(
			0,
			{ ranges: [], type: 'whole', operator: 'greaterThan', formula1: '0' },
			R('A5'),
		);
		const [list, whole] = ws(wb).dataValidations;
		expect(list?.ranges).toEqual([R('A1:A4'), R('A6:A10')]);
		expect(whole?.ranges).toEqual([R('A5')]);
		s.setDataValidation(0, undefined, R('A1:A10'));
		expect(ws(wb).dataValidations).toEqual([]);
	});
});

describe('view and names', () => {
	it('freezes and unfreezes panes', () => {
		const { wb, s } = setup();
		s.setFreeze(0, { rows: 1, cols: 2 });
		expect(ws(wb).view.freeze).toEqual({ rows: 1, cols: 2 });
		s.setFreeze(0, { rows: 0, cols: 0 });
		expect(ws(wb).view.freeze).toBeUndefined();
		s.undo();
		expect(ws(wb).view.freeze).toEqual({ rows: 1, cols: 2 });
	});
	it('patches the sheet view and clamps the zoom', () => {
		const { wb, s } = setup();
		s.setSheetView(0, { showGridLines: false, zoom: 1000 });
		expect(ws(wb).view.showGridLines).toBe(false);
		expect(ws(wb).view.zoom).toBe(400);
		s.undo();
		expect(ws(wb).view.zoom).toBe(100);
	});
	it('adds, replaces and deletes defined names', () => {
		const { wb, s } = setup();
		s.setDefinedName({ name: 'Rate', formula: '=0.2' });
		s.setDefinedName({ name: 'rate', formula: '0.25' });
		s.setDefinedName({ name: 'Rate', formula: 'Sheet2!$A$1', localSheet: 1 });
		expect(wb.definedNames).toEqual([
			{ name: 'rate', formula: '0.25' },
			{ name: 'Rate', formula: 'Sheet2!$A$1', localSheet: 1 },
		]);
		s.deleteDefinedName('RATE');
		expect(wb.definedNames).toHaveLength(1);
		s.undo();
		expect(wb.definedNames).toHaveLength(2);
	});
	it('validates defined names', () => {
		expect(validateDefinedName('Sales_2024')).toBeUndefined();
		expect(validateDefinedName('A1')).toBeDefined();
		expect(validateDefinedName('my name')).toBeDefined();
		expect(validateDefinedName('1abc')).toBeDefined();
		expect(validateDefinedName('R1C1')).toBeDefined();
		const { s } = setup();
		expect(() => s.setDefinedName({ name: 'B2', formula: '1' })).toThrow();
	});
});

describe('tables', () => {
	it('makes unique header names', () => {
		expect(uniqueHeaders(['Name', '', 'name', 'Qty'])).toEqual(['Name', 'Column2', 'name2', 'Qty']);
	});
	it('creates a table from a range with headers', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			['Name', 2024],
			['x', 1],
		]);
		const table = s.createTable(0, R('A1:B2'), true);
		expect(table).toMatchObject({
			id: 1,
			name: 'Table1',
			displayName: 'Table1',
			range: R('A1:B2'),
			headerRow: true,
		});
		expect(table.columns.map((c) => c.name)).toEqual(['Name', '2024']);
		expect(cell(wb, 'B1')?.value).toBe('2024');
		expect(s.createTable(1, R('A1:A2'), true).name).toBe('Table2');
	});
	it('inserts a header row when the range has none', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			[1, 2],
			[3, 4],
		]);
		const table = s.createTable(0, R('A1:B2'), false);
		expect(table.range).toEqual(R('A1:B3'));
		expect(cell(wb, 'A1')?.value).toBe('Column1');
		expect(cell(wb, 'A2')?.value).toBe(1);
		s.undo();
		expect(cell(wb, 'A1')?.value).toBe(1);
		expect(ws(wb).tables).toEqual([]);
	});
	it('expands a single cell to its region and refuses overlaps', () => {
		const { s } = setup();
		s.setRangeValues(0, A('A1'), [
			['a', 'b'],
			[1, 2],
			[3, 4],
		]);
		expect(s.createTable(0, R('B2'), true).range).toEqual(R('A1:B3'));
		expect(() => s.createTable(0, R('B1:C2'), true)).toThrow(/overlap/);
	});
});
