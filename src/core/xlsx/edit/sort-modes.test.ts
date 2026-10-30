import { describe, expect, it } from 'vitest';
import { parseXml } from '../../xml/index';
import { parseAddress, parseRange } from '../address';
import { getCell, putCell } from '../cells';
import { internStyle } from '../styles';
import type { CellValue, Workbook, Worksheet } from '../model';
import { loadXlsx } from '../read/load';
import { saveXlsx } from '../write/save';
import { readSortState } from '../read/sort-state';
import { BUILT_IN_SORT_LISTS, parseSortList } from '../sort-lists';
import { sortStateXml } from '../write/sort-state';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { compareCellValues, sortRows } from './sort';

const R = (ref: string) => parseRange(ref)!;
const A = (ref: string) => parseAddress(ref)!;
const sheetOf = (wb: Workbook): Worksheet => wb.sheets[0]!;
const col = (wb: Workbook, c: number, from: number, to: number): CellValue[] =>
	Array.from(
		{ length: to - from + 1 },
		(_v, i) => getCell(sheetOf(wb), from + i, c)?.value ?? null,
	);
const row = (wb: Workbook, r: number, from: number, to: number): CellValue[] =>
	Array.from(
		{ length: to - from + 1 },
		(_v, i) => getCell(sheetOf(wb), r, from + i)?.value ?? null,
	);

function coloured() {
	const wb = createWorkbook();
	const base = wb.styles[0]!;
	const red = internStyle(wb, {
		...base,
		fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } },
	});
	const blue = internStyle(wb, { ...base, font: { color: { rgb: 'FF0000FF' } } });
	const sheet = sheetOf(wb);
	putCell(sheet, 0, 0, { value: 'H' });
	putCell(sheet, 1, 0, { value: 'a' });
	putCell(sheet, 2, 0, { value: 'b', styleId: red });
	putCell(sheet, 3, 0, { value: 'c', styleId: blue });
	putCell(sheet, 4, 0, { value: 'd', styleId: red });
	return wb;
}

describe('sort on cell colour, font colour and icon', () => {
	it('puts a chosen fill on top, keeping the other rows in order', () => {
		const wb = coloured();
		sortRows(sheetOf(wb), R('A1:A5'), [{ col: 0, sortOn: 'cellColor', color: '#FF0000' }], true, {
			workbook: wb,
		});
		expect(col(wb, 0, 1, 4)).toEqual(['b', 'd', 'a', 'c']);
	});

	it('puts the chosen fill at the bottom when descending', () => {
		const wb = coloured();
		sortRows(
			sheetOf(wb),
			R('A1:A5'),
			[{ col: 0, sortOn: 'cellColor', color: '#FF0000', descending: true }],
			true,
			{ workbook: wb },
		);
		expect(col(wb, 0, 1, 4)).toEqual(['a', 'c', 'b', 'd']);
	});

	it('sorts "No Cell Color" first', () => {
		const wb = coloured();
		sortRows(sheetOf(wb), R('A1:A5'), [{ col: 0, sortOn: 'cellColor' }], true, { workbook: wb });
		expect(col(wb, 0, 1, 4)).toEqual(['a', 'c', 'b', 'd']);
	});

	it('sorts on font colour', () => {
		const wb = coloured();
		sortRows(sheetOf(wb), R('A1:A5'), [{ col: 0, sortOn: 'fontColor', color: '#0000FF' }], true, {
			workbook: wb,
		});
		expect(col(wb, 0, 1, 4)).toEqual(['c', 'a', 'b', 'd']);
	});

	it('sorts on a conditional-format icon and on its absence', () => {
		const wb = createWorkbook();
		const sheet = sheetOf(wb);
		['H', 'x', 'y', 'z'].forEach((v, i) => putCell(sheet, i, 0, { value: v }));
		const icons: Record<number, number | undefined> = { 1: 0, 2: undefined, 3: 2 };
		const formatAt = (r: number) => {
			const index = icons[r];
			return index === undefined ? undefined : { icon: { set: '3Arrows', index } as never };
		};
		sortRows(
			sheet,
			R('A1:A4'),
			[{ col: 0, sortOn: 'icon', icon: { set: '3Arrows', index: 2 } }],
			true,
			{ workbook: wb, formatAt },
		);
		expect(col(wb, 0, 1, 3)).toEqual(['z', 'x', 'y']);
		const wb2 = createWorkbook();
		['H', 'x', 'y', 'z'].forEach((v, i) => putCell(sheetOf(wb2), i, 0, { value: v }));
		sortRows(
			sheetOf(wb2),
			R('A1:A4'),
			[{ col: 0, sortOn: 'icon', icon: { set: '3Arrows' } }],
			true,
			{ workbook: wb2, formatAt },
		);
		expect(col(wb2, 0, 1, 3)).toEqual(['y', 'x', 'z']);
	});
});

describe('custom lists and case sensitivity', () => {
	it('orders by the built-in weekday list, unlisted values after', () => {
		const wb = createWorkbook();
		['H', 'Wed', 'Mon', 'Foo', 'Sun'].forEach((v, i) => putCell(sheetOf(wb), i, 0, { value: v }));
		const customList = BUILT_IN_SORT_LISTS[0]!;
		sortRows(sheetOf(wb), R('A1:A5'), [{ col: 0, customList }], true);
		expect(col(wb, 0, 1, 4)).toEqual(['Sun', 'Mon', 'Wed', 'Foo']);
		sortRows(sheetOf(wb), R('A1:A5'), [{ col: 0, customList, descending: true }], true);
		expect(col(wb, 0, 1, 4)).toEqual(['Wed', 'Mon', 'Sun', 'Foo']);
	});

	it('ranks a user list in the order typed', () => {
		const list = parseSortList('High, Medium\nLow');
		expect(list).toEqual(['High', 'Medium', 'Low']);
		expect(compareCellValues('Low', 'High', false, { customList: list })).toBeGreaterThan(0);
		expect(compareCellValues('low', 'High', false, { customList: list, caseSensitive: true })).toBe(
			1,
		);
	});

	it('puts lower case before upper case when case sensitive', () => {
		expect(compareCellValues('a', 'A')).toBe(0);
		expect(compareCellValues('a', 'A', false, { caseSensitive: true })).toBeLessThan(0);
		expect(compareCellValues('b', 'A', false, { caseSensitive: true })).toBeGreaterThan(0);
		const wb = createWorkbook();
		['H', 'B', 'b', 'A', 'a'].forEach((v, i) => putCell(sheetOf(wb), i, 0, { value: v }));
		sortRows(sheetOf(wb), R('A1:A5'), [{ col: 0 }], true, { caseSensitive: true });
		expect(col(wb, 0, 1, 4)).toEqual(['a', 'A', 'b', 'B']);
	});
});

describe('left to right', () => {
	it('reorders columns by a row, keeping the label column and formulas', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setRangeValues(0, A('A1'), [
			['Label', 3, 1, 2],
			['x', 'c', 'a', 'b'],
		]);
		s.sort(0, R('A1:D2'), [{ col: 0 }], true, { leftToRight: true });
		expect(row(wb, 0, 0, 3)).toEqual(['Label', 1, 2, 3]);
		expect(row(wb, 1, 0, 3)).toEqual(['x', 'a', 'b', 'c']);
	});

	it('shifts relative references by the column move', () => {
		const wb = createWorkbook();
		const sheet = sheetOf(wb);
		putCell(sheet, 0, 0, { value: 9 });
		putCell(sheet, 0, 1, { value: 1 });
		putCell(sheet, 1, 0, { value: 100 });
		putCell(sheet, 1, 1, { value: 200 });
		putCell(sheet, 2, 0, { value: 0, formula: 'A2*2' });
		sortRows(sheet, R('A1:B3'), [{ col: 0 }], false, { leftToRight: true });
		expect(row(wb, 0, 0, 1)).toEqual([1, 9]);
		expect(getCell(sheet, 2, 1)?.formula).toBe('B2*2');
	});
});

describe('sort state', () => {
	it('records every level in the sheet and round-trips through the file XML', () => {
		const wb = coloured();
		const s = createEditSession(wb, { recalc: false });
		s.sort(
			0,
			R('A1:A5'),
			[
				{ col: 0, sortOn: 'cellColor', color: '#FF0000', descending: true },
				{ col: 0, customList: ['a', 'b'] },
			],
			true,
			{ caseSensitive: true },
		);
		const state = sheetOf(wb).sortState!;
		expect(state.range).toEqual(R('A2:A5'));
		expect(state.caseSensitive).toBe(true);
		expect(state.conditions[0]).toMatchObject({ sortBy: 'cellColor', descending: true });
		expect(state.conditions[1]?.customList).toBe('a,b');
		const xml = sortStateXml(state);
		expect(xml).toContain('caseSensitive="1"');
		expect(xml).toContain('sortBy="cellColor"');
		const again = readSortState(
			parseXml(
				xml.replace(
					'<sortState',
					'<sortState xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
				),
			).documentElement,
		);
		expect(again?.conditions.map((c) => [c.sortBy, c.descending, c.customList])).toEqual([
			['cellColor', true, undefined],
			[undefined, undefined, 'a,b'],
		]);
	});

	it('survives a save and load, colour format included', async () => {
		const wb = coloured();
		createEditSession(wb, { recalc: false }).sort(
			0,
			R('A1:A5'),
			[{ col: 0, sortOn: 'cellColor', color: '#FF0000' }],
			true,
		);
		const loaded = await loadXlsx(await saveXlsx(wb));
		const state = loaded.sheets[0]!.sortState!;
		expect(state.range).toEqual(R('A2:A5'));
		expect(state.conditions[0]?.sortBy).toBe('cellColor');
		expect(state.conditions[0]?.style?.fill).toBeDefined();
		const again = await loadXlsx(await saveXlsx(loaded));
		expect(again.sheets[0]!.sortState).toEqual(state);
	});

	it('reads and writes dxf, icon and column-sort attributes', () => {
		const source =
			'<sortState xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" columnSort="1" ref="B1:D2">' +
			'<sortCondition sortBy="icon" ref="B1:D1" iconSet="3Arrows" iconId="1"/>' +
			'<sortCondition descending="1" sortBy="fontColor" ref="B2:D2" dxfId="3"/></sortState>';
		const state = readSortState(parseXml(source).documentElement)!;
		expect(state.columnSort).toBe(true);
		expect(state.conditions[0]).toMatchObject({ sortBy: 'icon', iconSet: '3Arrows', iconId: 1 });
		expect(state.conditions[1]).toMatchObject({ sortBy: 'fontColor', descending: true, dxfId: 3 });
		const out = sortStateXml(state);
		expect(out).toContain('columnSort="1"');
		expect(out).toContain('iconSet="3Arrows"');
		expect(out).toContain('dxfId="3"');
	});
});
