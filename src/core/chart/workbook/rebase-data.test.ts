import { describe, expect, it } from 'vitest';
import { columnChart } from '../edit/test-support/column-chart';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { rebaseChartData } from './rebase-data';

describe('fresh classic chart data ranges', () => {
	it('allocates names, caches and blanks without mutating the original references', () => {
		const source = columnChart();
		const series = source.plotArea.groups[0]!.series[0]!;
		series.values!.cache = {
			type: 'number',
			pointCount: 3,
			points: [
				{ index: 0, value: '2.5' },
				{ index: 2, value: '0' },
			],
		};
		const original = structuredClone(source);
		const result = rebaseChartData(source, "Chart's data");
		const first = result.chartSpace.plotArea.groups[0]!.series[0]!;
		expect(first.values!.formula).toBe("'Chart''s data'!$B$2:$B$4");
		expect(first.tx!.reference!.formula).toBe("'Chart''s data'!$B$1");
		expect(result.rows.slice(1).map((row) => row[1])).toEqual([2.5, null, 0]);
		expect(source).toEqual(original);
		expect(
			parseChartSpace(writeChartSpace(result.chartSpace)).chartSpace.plotArea.groups[0]!.series[0]!
				.values,
		).toEqual(first.values);
	});
	it('writes hierarchy columns outermost first and keeps per-level cache order', () => {
		const source = columnChart();
		source.plotArea.groups[0]!.series[0]!.categories = {
			kind: 'multiLvlStrRef',
			levelPointCount: 2,
			levels: [
				{
					type: 'string',
					pointCount: 2,
					points: [
						{ index: 0, value: 'Leaf A' },
						{ index: 1, value: 'Leaf B' },
					],
				},
				{ type: 'string', pointCount: 2, points: [{ index: 0, value: 'Parent' }] },
			],
		};
		const result = rebaseChartData(source, 'NewData');
		expect(result.rows[1]!.slice(0, 2)).toEqual(['Parent', 'Leaf A']);
		expect(result.chartSpace.plotArea.groups[0]!.series[0]!.categories!.formula).toBe(
			'NewData!$A$2:$B$3',
		);
		expect(result.chartSpace.plotArea.groups[0]!.series[0]!.categories!.levels).toEqual(
			source.plotArea.groups[0]!.series[0]!.categories!.levels,
		);
	});
	it('refuses an external source and missing caches', () => {
		const source = columnChart();
		source.plotArea.groups[0]!.series[0]!.values!.formula = '[External.xlsx]Sheet1!A1:A2';
		expect(() => rebaseChartData(source, 'NewData')).toThrow('External');
		delete source.plotArea.groups[0]!.series[0]!.values!.formula;
		delete source.plotArea.groups[0]!.series[0]!.values!.cache;
		expect(() => rebaseChartData(source, 'NewData')).toThrow('cached');
	});
});
