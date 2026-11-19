import { describe, expect, it } from 'vitest';

import {
	CHART_POINTS_LIMITED,
	MAX_CHART_BLANK_POINTS,
	chartPointLimit,
	chartPointLimitOfXml,
} from '../../../chart/point-limit';
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
	it('reports an oversized part and builds a bounded series from it', () => {
		const { chartSpace, pointLimitIssue } = readChartPartModel(part(HUGE));
		expect(pointLimitIssue?.code).toBe(CHART_POINTS_LIMITED);
		const values = neutralNumbersWithBlanks(chartSpace!.plotArea.groups[0]!.series[0]!.values);
		expect(values[0]).toBe(1);
		expect(values.length).toBeLessThanOrEqual(MAX_CHART_BLANK_POINTS + 1);
	});

	it('leaves a sparse part alone', () => {
		const { chartSpace, pointLimitIssue } = readChartPartModel(part(4, 2));
		expect(pointLimitIssue).toBeUndefined();
		expect(neutralNumbersWithBlanks(chartSpace!.plotArea.groups[0]!.series[0]!.values)).toEqual([
			null,
			null,
			1,
			null,
		]);
	});

	it('counts the caches of a part from its text', () => {
		const xml = `${part(1)}<c15:dlblRangeCache/><cx:lvl ptCount="1"/>`;
		expect(chartPointLimitOfXml(xml).blanks).toBe(Math.floor(MAX_CHART_BLANK_POINTS / 3));
	});

	it('bounds the object-tree readers', () => {
		const numeric = { 'c:numRef': { 'c:numCache': cache(HUGE, [0, '1']) } };
		expect(extractSeriesNumbersWithBlanks(numeric, lookup).length).toBeLessThanOrEqual(
			MAX_CHART_BLANK_POINTS + 1,
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
		expect(categories.length).toBeLessThanOrEqual(2 * MAX_CHART_BLANK_POINTS + 1);
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
		const limit = chartPointLimit(4);
		const [series] = parseFilteredSeries(container, lookup, limit)!;
		expect(series!.categories![0]).toBe('a');
		expect(series!.categories!.length).toBeLessThanOrEqual(limit.blanks + 1);
		expect(limit.limited).toBe(true);
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
		const limit = chartPointLimit(2);
		const range = parseDataLabelsRange(group, lookup, limit)!;
		expect(range.cache.slice(0, 2)).toEqual(['', 'x']);
		expect(range.cache.length).toBeLessThanOrEqual(limit.blanks + 2);
		expect(limit.limited).toBe(true);
	});
});
