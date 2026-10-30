import { describe, expect, it } from 'vitest';
import { parseRange } from './address';
import { putCell } from './cells';
import { compileAutoFilter } from './filter-criteria';
import { dateToSerial } from './numfmt/date';
import type { AutoFilterColumn, CellValue } from './model';
import { createWorkbook } from './workbook';

function visible(values: CellValue[], column: Omit<AutoFilterColumn, 'offset'>, today?: number) {
	const workbook = createWorkbook();
	const sheet = workbook.sheets[0]!;
	putCell(sheet, 0, 0, { value: 'H' });
	values.forEach((value, i) => putCell(sheet, i + 1, 0, { value }));
	const filter = {
		range: parseRange(`A1:A${values.length + 1}`)!,
		columns: [{ offset: 0, ...column }],
	};
	const test = compileAutoFilter(workbook, sheet, filter, today === undefined ? {} : { today });
	return values.filter((_, i) => test(i + 1));
}

describe('filter criteria evaluation', () => {
	it('matches the value checklist and blanks', () => {
		expect(visible(['a', 'b', null], { values: ['a'], blank: true })).toEqual(['a', null]);
		expect(visible(['a', 'b'], { values: [] })).toEqual([]);
	});

	it('matches date groups', () => {
		const d = (y: number, m: number, day: number) =>
			dateToSerial(new Date(Date.UTC(y, m - 1, day)));
		const values = [d(2024, 3, 5), d(2024, 4, 1), d(2023, 3, 5)];
		expect(
			visible(values, { values: [], dateGroups: [{ year: 2024, month: 3, grouping: 'month' }] }),
		).toEqual([values[0]]);
		expect(visible(values, { values: [], dateGroups: [{ year: 2024, grouping: 'year' }] })).toEqual(
			[values[0], values[1]],
		);
	});

	it('evaluates number comparisons joined by and / or', () => {
		const rows = [1, 5, 10, 20];
		const gt = { operator: 'greaterThan', value: '4' } as const;
		const lt = { operator: 'lessThan', value: '15' } as const;
		expect(visible(rows, { custom: { and: true, conditions: [gt, lt] } })).toEqual([5, 10]);
		expect(
			visible(rows, {
				custom: {
					conditions: [
						{ operator: 'lessThan', value: '2' },
						{ operator: 'greaterThan', value: '15' },
					],
				},
			}),
		).toEqual([1, 20]);
		expect(
			visible(rows, { custom: { conditions: [{ operator: 'notEqual', value: '5' }] } }),
		).toEqual([1, 10, 20]);
	});

	it('evaluates text conditions with wildcards, case-insensitively', () => {
		const rows = ['Apple', 'apricot', 'Banana', null];
		const eq = (value: string, operator: 'equal' | 'notEqual' = 'equal') => ({
			custom: { conditions: [{ operator, value }] },
		});
		expect(visible(rows, eq('ap*'))).toEqual(['Apple', 'apricot']);
		expect(visible(rows, eq('*an*'))).toEqual(['Banana']);
		expect(visible(rows, eq('?pple'))).toEqual(['Apple']);
		expect(visible(rows, eq('*an*', 'notEqual'))).toEqual(['Apple', 'apricot', null]);
		expect(visible(rows, eq('BANANA'))).toEqual(['Banana']);
		expect(visible(['a*b', 'axb'], eq('a~*b'))).toEqual(['a*b']);
	});

	it('treats notEqual with a single space as "non-blanks"', () => {
		expect(
			visible(['a', null, 3], { custom: { conditions: [{ operator: 'notEqual', value: ' ' }] } }),
		).toEqual(['a', 3]);
	});

	it('keeps the top and bottom items or percent, ties included', () => {
		const rows = [5, 9, 9, 1, 7, 3, 'x'];
		expect(visible(rows, { top10: { top: true, percent: false, value: 2 } })).toEqual([9, 9]);
		expect(visible(rows, { top10: { top: false, percent: false, value: 1 } })).toEqual([1]);
		const ten = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
		expect(visible(ten, { top10: { top: true, percent: true, value: 20 } })).toEqual([9, 10]);
	});

	it('filters above and below average', () => {
		const rows = [1, 2, 3, 10];
		expect(visible(rows, { dynamic: { type: 'aboveAverage' } })).toEqual([10]);
		expect(visible(rows, { dynamic: { type: 'belowAverage' } })).toEqual([1, 2, 3]);
	});

	it('filters date periods against today', () => {
		const d = (y: number, m: number, day: number) =>
			dateToSerial(new Date(Date.UTC(y, m - 1, day)));
		const today = d(2024, 5, 15);
		const rows = [d(2024, 5, 15), d(2024, 5, 14), d(2024, 4, 30), d(2023, 5, 15), d(2024, 12, 1)];
		expect(visible(rows, { dynamic: { type: 'today' } }, today)).toEqual([rows[0]]);
		expect(visible(rows, { dynamic: { type: 'thisMonth' } }, today)).toEqual([rows[0], rows[1]]);
		expect(visible(rows, { dynamic: { type: 'lastMonth' } }, today)).toEqual([rows[2]]);
		expect(visible(rows, { dynamic: { type: 'lastYear' } }, today)).toEqual([rows[3]]);
		expect(visible(rows, { dynamic: { type: 'thisQuarter' } }, today)).toEqual([
			rows[0],
			rows[1],
			rows[2],
		]);
		expect(visible(rows, { dynamic: { type: 'yearToDate' } }, today)).toEqual([
			rows[0],
			rows[1],
			rows[2],
		]);
		expect(visible(rows, { dynamic: { type: 'M12' } }, today)).toEqual([rows[4]]);
		expect(visible(rows, { dynamic: { type: 'Q2' } }, today)).toEqual([
			rows[0],
			rows[1],
			rows[2],
			rows[3],
		]);
	});

	it('filters by cell fill and font colour', () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		workbook.styles.push(
			{
				...workbook.styles[0]!,
				fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } },
			},
			{ ...workbook.styles[0]!, font: { color: { rgb: 'FF0000FF' } } },
		);
		putCell(sheet, 0, 0, { value: 'H' });
		putCell(sheet, 1, 0, { value: 'red', styleId: workbook.styles.length - 2 });
		putCell(sheet, 2, 0, { value: 'blue text', styleId: workbook.styles.length - 1 });
		putCell(sheet, 3, 0, { value: 'plain' });
		const range = parseRange('A1:A4')!;
		const run = (color: NonNullable<AutoFilterColumn['color']>) => {
			const test = compileAutoFilter(workbook, sheet, { range, columns: [{ offset: 0, color }] });
			return [1, 2, 3].filter(test);
		};
		const red = {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } },
		} as const;
		expect(run({ cellColor: true, style: red })).toEqual([1]);
		expect(run({ cellColor: true, style: { fill: { type: 'pattern', pattern: 'none' } } })).toEqual(
			[2, 3],
		);
		expect(run({ cellColor: false, style: { font: { color: { rgb: 'FF0000FF' } } } })).toEqual([2]);
		expect(run({ cellColor: false, style: {} })).toEqual([1, 3]);
	});

	it('uses the conditional-format icon when supplied', () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		for (let i = 0; i < 3; i++) putCell(sheet, i + 1, 0, { value: i });
		const range = parseRange('A1:A4')!;
		const test = compileAutoFilter(
			workbook,
			sheet,
			{ range, columns: [{ offset: 0, icon: { iconSet: '3Arrows', iconId: 2 } }] },
			{ formatAt: (row) => ({ icon: { set: '3Arrows', index: row - 1 } }) },
		);
		expect([1, 2, 3].filter(test)).toEqual([3]);
	});
});
