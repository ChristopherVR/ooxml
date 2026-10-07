import type { PptxChartData } from '../types/chart';
import { describe, expect, it } from 'vitest';

import {
	chartGridAddCategory as addChartCategory,
	chartGridAddSeries as addChartSeries,
	chartGridRemoveCategory as removeChartCategory,
	chartGridRemoveSeries as removeChartSeries,
	chartGridSetCategoryLabel as setChartCategoryLabel,
	chartGridSetCellValue as setChartCellValue,
} from './chart-grid-operations';

function makeData(): PptxChartData {
	return {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [
			{ name: 'Revenue', values: [10, 20] },
			{ name: 'Cost', values: [5, 6] },
		],
	};
}

describe('addChartSeries', () => {
	it('appends a zero-filled series named after its position', () => {
		const next = addChartSeries(makeData());
		expect(next.series).toHaveLength(3);
		expect(next.series[2].name).toBe('Series 3');
		expect(next.series[2].values).toStrictEqual([0, 0]);
	});
});

describe('removeChartSeries', () => {
	it('removes the addressed series', () => {
		const next = removeChartSeries(makeData(), 0);
		expect(next?.series.map((s) => s.name)).toStrictEqual(['Cost']);
	});

	it('refuses to remove the last series', () => {
		const single: PptxChartData = { ...makeData(), series: [{ name: 'Only', values: [1, 2] }] };
		expect(removeChartSeries(single, 0)).toBeNull();
	});

	it('refuses an out-of-range index', () => {
		expect(removeChartSeries(makeData(), 9)).toBeNull();
	});
});

describe('addChartCategory', () => {
	it('appends a category and pads every series', () => {
		const next = addChartCategory(makeData());
		expect(next.categories).toStrictEqual(['Q1', 'Q2', 'Cat 3']);
		expect(next.series[0].values).toHaveLength(3);
		expect(next.series[1].values).toHaveLength(3);
	});
});

describe('removeChartCategory', () => {
	it('removes the category from the labels and every series', () => {
		const next = removeChartCategory(makeData(), 0);
		expect(next?.categories).toStrictEqual(['Q2']);
		expect(next?.series[0].values).toStrictEqual([20]);
	});

	it('refuses to remove the last category', () => {
		const single: PptxChartData = {
			...makeData(),
			categories: ['Only'],
			series: [{ name: 'Revenue', values: [1] }],
		};
		expect(removeChartCategory(single, 0)).toBeNull();
	});
});

describe('setChartCategoryLabel', () => {
	it('renames only the addressed label', () => {
		expect(setChartCategoryLabel(makeData(), 1, 'Second')?.categories).toStrictEqual([
			'Q1',
			'Second',
		]);
	});

	it('refuses an out-of-range index', () => {
		expect(setChartCategoryLabel(makeData(), 4, 'x')).toBeNull();
	});
});

describe('setChartCellValue', () => {
	it('writes a parsed numeric cell', () => {
		expect(setChartCellValue(makeData(), 0, 1, '42')?.series[0].values).toStrictEqual([10, 42]);
	});

	it('rejects a cleared cell rather than coercing it to zero', () => {
		expect(setChartCellValue(makeData(), 0, 1, '')).toBeNull();
		expect(setChartCellValue(makeData(), 0, 1, 'abc')).toBeNull();
	});

	it('refuses an out-of-range series', () => {
		expect(setChartCellValue(makeData(), 7, 0, '1')).toBeNull();
	});
});

describe('grid editing guards and preservation', () => {
	it.each([true, false])('retains point-formatting policy: followDataPoint=%s', (follow) => {
		const data = makeData();
		data.series[0].dataPoints = [{ idx: 1 }];
		const next = removeChartCategory(data, 0, follow);
		expect(next?.series[0].dataPoints).toEqual([{ idx: follow ? 0 : 1 }]);
		expect(data.series[0].dataPoints).toEqual([{ idx: 1 }]);
	});

	it.each([-1, 2, 0.5, Number.NaN, Number.POSITIVE_INFINITY])(
		'rejects invalid grid indices: %s',
		(index) => {
			expect(removeChartSeries(makeData(), index)).toBeNull();
			expect(removeChartCategory(makeData(), index)).toBeNull();
			expect(setChartCategoryLabel(makeData(), index, 'x')).toBeNull();
			expect(setChartCellValue(makeData(), index, 0, '42')).toBeNull();
			expect(setChartCellValue(makeData(), 0, index, '42')).toBeNull();
		},
	);

	it('preserves the source when editing a cell', () => {
		const data = makeData();
		const before = structuredClone(data);
		const next = setChartCellValue(data, 0, 1, '42');
		expect(data).toEqual(before);
		expect(next?.series[0].values).toEqual([10, 42]);
		expect(next?.series[1]).toEqual(data.series[1]);
		expect(next?.chartType).toBe('bar');
	});
});
