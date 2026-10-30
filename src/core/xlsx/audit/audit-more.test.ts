import { describe, expect, it } from 'vitest';
import { book, engine, locate } from '../formula/test-helpers';
import { findErrors } from './error-check';
import { directDependents, directPrecedents, traceNext } from './trace';

const at = (wb: ReturnType<typeof book>, key: string) => locate(wb, key);
const table = (name: string, endRow: number, endCol: number, columns: string[]) => ({
	id: 1,
	name,
	displayName: name,
	range: { start: { row: 0, col: 0 }, end: { row: endRow, col: endCol } },
	headerRow: false,
	totalsRow: false,
	columns: columns.map((c) => ({ name: c })),
});

describe('trace through names, tables and 3D references', () => {
	it('resolves them with the engine', () => {
		const wb = book({
			A1: 1,
			A2: 2,
			'Sheet2!A1': 5,
			B1: '=Total*2',
			B2: '=SUM(Sheet1:Sheet2!A1)',
			B3: '=SUM(Tbl[Amt])',
		});
		wb.definedNames.push({ name: 'Total', formula: 'Sheet1!$A$1:$A$2' });
		wb.sheets[0]!.tables.push(table('Tbl', 1, 0, ['Amt']));
		const calc = engine(wb);
		const source = (c: { sheet: number; row: number; col: number }) => calc.precedents(c);
		expect(directPrecedents(wb, at(wb, 'B1'), source).map((p) => p.range.end.row)).toEqual([1]);
		expect(
			directPrecedents(wb, at(wb, 'B2'), source)
				.map((p) => p.sheet)
				.sort(),
		).toEqual([0, 1]);
		expect(directPrecedents(wb, at(wb, 'B3'), source)).toHaveLength(1);
		const t = traceNext(wb, [], 'precedents', at(wb, 'B2'), source);
		expect(t.arrows.some((a) => a.offSheet)).toBe(true);
		expect(directDependents(wb, at(wb, 'A1'), source).map((d) => d.row)).toEqual([0, 1, 2]);
	});
});

describe('more error rules', () => {
	it('flags references to empty cells only when asked', () => {
		const wb = book({ A1: '=Z9+1' });
		expect(findErrors(wb, 0)).toEqual([]);
		expect(findErrors(wb, 0, ['empty-cell-reference']).map((i) => i.rule)).toEqual([
			'empty-cell-reference',
		]);
	});

	it('flags a calculated column cell that differs from its column', () => {
		const wb = book({
			A1: 1,
			A2: 2,
			A3: 3,
			A4: 4,
			B1: '=A1*2',
			B2: '=A2*2',
			B3: '=A3*3',
			B4: '=A4*2',
		});
		wb.sheets[0]!.tables.push(table('T', 3, 1, ['A', 'B']));
		expect(findErrors(wb, 0, ['calculated-column']).map((i) => [i.row, i.col, i.rule])).toEqual([
			[2, 1, 'calculated-column'],
		]);
	});
});
