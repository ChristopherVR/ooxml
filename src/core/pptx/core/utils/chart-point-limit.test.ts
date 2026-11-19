import { describe, expect, it } from 'vitest';

import { MAX_CACHE_BLANK_POINTS, chartPointBudget } from '../../../chart/point-limit';
import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { XmlObject } from '../types';
import { extractSeriesNumbersWithBlanks } from './chart-blank-values';
import { CHART15_EXT_URI } from './chart-data-label-field-table';
import { parseDataLabelsRange } from './chart-data-labels-range';
import { CHART_FILTERED_SERIES_EXT_URI, parseFilteredSeries } from './chart-filtered-series';
import { neutralNumbersWithBlanks } from './chart-from-neutral-data';
import { extractMultiLevelCategoryValues } from './chart-multi-level-category-parser';
import { readChartPartModel } from './chart-part-repair';

const HUGE = 2_000_000_000;
const lookup = new PptxXmlLookupService();

const cache = (count: number, ...points: [number, string][]): XmlObject => ({
	'c:ptCount': { '@_val': String(count) },
	'c:pt': points.map(([idx, value]) => ({ '@_idx': String(idx), 'c:v': value })),
});

const part = (count: number, index = 0) =>
	`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>S!$A$1</c:f><c:numCache><c:ptCount val="${count}"/><c:pt idx="${index}"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;

describe('PowerPoint chart readers and declared point counts', () => {
	it('keeps the declared count of an oversized part and builds a bounded series from it', () => {
		const { chartSpace } = readChartPartModel(part(HUGE));
		const source = chartSpace!.plotArea.groups[0]!.series[0]!.values!;
		expect(source.cache!.pointCount).toBe(HUGE);
		expect(neutralNumbersWithBlanks(source)).toEqual([1]);
	});

	it('leaves a sparse part alone', () => {
		const { chartSpace } = readChartPartModel(part(4, 2));
		expect(neutralNumbersWithBlanks(chartSpace!.plotArea.groups[0]!.series[0]!.values)).toEqual([
			null,
			null,
			1,
			null,
		]);
	});

	it('spends one budget across the parts of a presentation', () => {
		const budget = chartPointBudget();
		const lengths = Array.from({ length: 6 }, () => {
			const { chartSpace } = readChartPartModel(part(MAX_CACHE_BLANK_POINTS), budget);
			return neutralNumbersWithBlanks(chartSpace!.plotArea.groups[0]!.series[0]!.values).length;
		});
		expect(lengths).toEqual([
			MAX_CACHE_BLANK_POINTS,
			MAX_CACHE_BLANK_POINTS,
			MAX_CACHE_BLANK_POINTS,
			MAX_CACHE_BLANK_POINTS,
			1,
			1,
		]);
	});

	it('bounds the object-tree readers', () => {
		const numeric = { 'c:numRef': { 'c:numCache': cache(HUGE, [0, '1']) } };
		expect(extractSeriesNumbersWithBlanks(numeric, lookup).length).toBeLessThanOrEqual(
			MAX_CACHE_BLANK_POINTS + 1,
		);
		const levels = {
			'c:multiLvlStrRef': {
				'c:multiLvlStrCache': {
					'c:ptCount': { '@_val': String(HUGE) },
					'c:lvl': { 'c:pt': { '@_idx': String(HUGE), 'c:v': 'a' } },
				},
			},
		};
		const categories = extractMultiLevelCategoryValues(levels, lookup)!.categories;
		expect(categories.length).toBeLessThanOrEqual(MAX_CACHE_BLANK_POINTS + 1);
	});

	it('bounds hidden series and records the cut', () => {
		const container = {
			'c:extLst': {
				'c:ext': {
					'@_uri': CHART_FILTERED_SERIES_EXT_URI,
					'c15:filteredBarSeries': {
						'c15:ser': {
							'c:idx': { '@_val': '1' },
							'c:cat': { 'c:strRef': { 'c:strCache': cache(HUGE, [0, 'a']) } },
						},
					},
				},
			},
		};
		const budget = chartPointBudget();
		const [series] = parseFilteredSeries(container, lookup, budget)!;
		expect(series!.categories).toEqual(['a']);
		expect(budget.cuts).toBe(1);
	});

	it('bounds a label range cache and records the cut', () => {
		const group = {
			'c:extLst': {
				'c:ext': {
					'@_uri': CHART15_EXT_URI,
					'c15:datalabelsRange': {
						'c15:f': 'Sheet1!$D$2:$D$4',
						'c15:dlblRangeCache': cache(HUGE, [1, 'x']),
					},
				},
			},
		};
		const budget = chartPointBudget();
		const range = parseDataLabelsRange(group, lookup, budget)!;
		expect(range.cache).toEqual(['', 'x']);
		expect(budget.cuts).toBe(1);
	});
});
