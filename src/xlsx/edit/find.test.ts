import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { queryPattern, replaceText } from './find.js';
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
const cell = (wb: Workbook, ref: string, i = 0) => {
	const ws = wb.sheets[i];
	return ws && getCell(ws, A(ref).row, A(ref).col);
};
const setup = () => {
	const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
	const s = createEditSession(wb, { recalc: false });
	s.setRangeValues(0, A('A1'), [
		['Apple pie', 'apple'],
		['Banana', 'APPLE'],
	]);
	s.setCellInput(0, 2, 0, '=SUM(A1:A2)');
	s.setCellInput(0, 2, 1, '0.5');
	s.applyStyle(0, [R('B3')], { numFmt: '0%' });
	s.setCellValue(1, 0, 0, 'pineapple');
	return { wb, s };
};
const refs = (matches: { sheet: number; row: number; col: number }[]) =>
	matches.map((m) => `${m.sheet}:${String.fromCharCode(65 + m.col)}${m.row + 1}`);

describe('queryPattern', () => {
	it('supports wildcards and escapes', () => {
		expect(queryPattern({ text: 'a*e' })?.test('apple')).toBe(true);
		expect(queryPattern({ text: 'a?e', wholeCell: true })?.test('ape')).toBe(true);
		expect(queryPattern({ text: 'a?e', wholeCell: true })?.test('apple')).toBe(false);
		expect(queryPattern({ text: '~*' })?.test('5*3')).toBe(true);
		expect(queryPattern({ text: '~*' })?.test('53')).toBe(false);
		expect(queryPattern({ text: 'a.b' })?.test('axb')).toBe(false);
		expect(queryPattern({ text: '*', wildcards: false })?.test('x')).toBe(false);
		expect(queryPattern({ text: '' })).toBeUndefined();
	});
	it('replaces every occurrence, case-insensitively by default', () => {
		expect(replaceText('Apple apple', { text: 'apple' }, 'pear')).toBe('pear pear');
		expect(replaceText('Apple apple', { text: 'apple', matchCase: true }, 'pear')).toBe(
			'Apple pear',
		);
		expect(replaceText('cost $5', { text: 'cost' }, '$&')).toBe('$& $5');
	});
});

describe('findAll', () => {
	it('finds case-insensitive matches on every sheet in row order', () => {
		const { s } = setup();
		expect(refs(s.findAll({ text: 'apple' }))).toEqual(['0:A1', '0:B1', '0:B2', '1:A1']);
	});
	it('matches case when asked', () => {
		const { s } = setup();
		expect(refs(s.findAll({ text: 'apple', matchCase: true }))).toEqual(['0:B1', '1:A1']);
	});
	it('matches the whole cell', () => {
		const { s } = setup();
		expect(refs(s.findAll({ text: 'apple', wholeCell: true }))).toEqual(['0:B1', '0:B2']);
	});
	it('searches formulas or values', () => {
		const { s } = setup();
		expect(refs(s.findAll({ text: 'SUM' }))).toEqual(['0:A3']);
		expect(refs(s.findAll({ text: 'SUM', lookIn: 'values' }))).toEqual([]);
		expect(refs(s.findAll({ text: '50%', lookIn: 'values' }))).toEqual(['0:B3']);
		expect(refs(s.findAll({ text: '0.5', lookIn: 'formulas' }))).toEqual(['0:B3']);
	});
	it('limits the search to a sheet and range, and orders by columns', () => {
		const { s } = setup();
		expect(refs(s.findAll({ text: 'a', sheet: 0, range: R('B1:B2') }))).toEqual(['0:B1', '0:B2']);
		expect(refs(s.findAll({ text: 'a', sheet: 0, order: 'columns' })).slice(0, 3)).toEqual([
			'0:A1',
			'0:A2',
			'0:A3',
		]);
	});
	it('searches comments', () => {
		const { s } = setup();
		s.setComment(0, A('D4'), 'check apple price', 'me');
		expect(refs(s.findAll({ text: 'price', lookIn: 'comments' }))).toEqual(['0:D4']);
	});
});

describe('replace', () => {
	it('replaces in all matching cells as one undo step', () => {
		const { wb, s } = setup();
		expect(s.replaceAll({ text: 'apple' }, 'pear')).toBe(4);
		expect(cell(wb, 'A1')?.value).toBe('pear pie');
		expect(cell(wb, 'B2')?.value).toBe('pear');
		expect(cell(wb, 'A1', 1)?.value).toBe('pinepear');
		s.undo();
		expect(cell(wb, 'A1')?.value).toBe('Apple pie');
		expect(cell(wb, 'A1', 1)?.value).toBe('pineapple');
	});
	it('replaces whole cells only', () => {
		const { wb, s } = setup();
		expect(s.replaceAll({ text: 'apple', wholeCell: true, sheet: 0 }, 'x')).toBe(2);
		expect(cell(wb, 'A1')?.value).toBe('Apple pie');
	});
	it('replaces inside formulas and re-parses the result', () => {
		const { wb, s } = setup();
		s.replaceAll({ text: 'SUM' }, 'MAX');
		expect(cell(wb, 'A3')?.formula).toBe('MAX(A1:A2)');
		s.setCellValue(0, 5, 0, 'n100');
		s.replaceAll({ text: 'n', wholeCell: false, sheet: 0, range: R('A6') }, '');
		expect(cell(wb, 'A6')?.value).toBe(100);
	});
	it('replaces one match', () => {
		const { wb, s } = setup();
		const [first] = s.findAll({ text: 'apple' });
		if (!first) throw new Error('no match');
		expect(s.replaceOne({ text: 'apple' }, 'kiwi', first)).toBe(true);
		expect(cell(wb, 'A1')?.value).toBe('kiwi pie');
		expect(s.replaceOne({ text: 'apple' }, 'kiwi', first)).toBe(false);
	});
	it('replaces in comments', () => {
		const { wb, s } = setup();
		s.setComment(0, A('A1'), 'old note', 'me');
		expect(s.replaceAll({ text: 'old', lookIn: 'comments' }, 'new')).toBe(1);
		expect(wb.sheets[0]?.comments[0]?.text).toBe('new note');
	});
	it('returns 0 when nothing matches', () => {
		const { s } = setup();
		expect(s.replaceAll({ text: 'zzz' }, 'y')).toBe(0);
		expect(s.undoLabel()).not.toBe('Replace');
	});
});
