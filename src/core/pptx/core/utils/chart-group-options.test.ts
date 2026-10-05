import { describe, expect, it } from 'vitest';

import type { PptxChartData, XmlObject } from '../types';
import {
	applyChartGroupOptionsToXml,
	chartGroupOptionVal,
	CHART_GROUP_OPTIONS,
} from './chart-group-options';

const localName = (key: string) => key.replace(/^.*:/u, '');

function data(overrides: Partial<PptxChartData>): PptxChartData {
	return { chartType: 'bar', categories: [], series: [], ...overrides };
}

function barContainer(): XmlObject {
	return {
		'c:barDir': { '@_val': 'bar' },
		'c:ser': {},
		'c:gapWidth': { '@_val': '35' },
		'c:overlap': { '@_val': '10' },
		'c:axId': [],
	};
}

describe('applyChartGroupOptionsToXml (loaded charts)', () => {
	it('leaves elements the model still describes untouched', () => {
		const container = barContainer();
		const gapWidth = container['c:gapWidth'];
		const changed = applyChartGroupOptionsToXml(
			container,
			'barChart',
			data({ barGapWidth: 35, barOverlap: 10 }),
			localName,
		);
		expect(changed).toBe(false);
		expect(container['c:gapWidth']).toBe(gapWidth);
	});

	it('rewrites an edited value in place and inserts a new one', () => {
		const container = barContainer();
		delete container['c:overlap'];
		applyChartGroupOptionsToXml(
			container,
			'barChart',
			data({ barGapWidth: 91, barOverlap: -4 }),
			localName,
		);
		expect(container['c:gapWidth']).toStrictEqual({ '@_val': '91' });
		expect(container['c:overlap']).toStrictEqual({ '@_val': '-4' });
	});

	it('removes a value the model cleared and ignores elements the container cannot hold', () => {
		const container = barContainer();
		applyChartGroupOptionsToXml(
			container,
			'barChart',
			data({ barGapWidth: 35, doughnutHoleSize: 40, firstSliceAngle: 90 }),
			localName,
		);
		expect(container['c:overlap']).toBeUndefined();
		expect(container['c:holeSize']).toBeUndefined();
		expect(container['c:firstSliceAng']).toBeUndefined();
	});

	it('keeps an authored value it cannot parse and leaves ofPie gapWidth to the of-pie writer', () => {
		const container: XmlObject = { 'c:gapWidth': { '@_val': 'oops' } };
		applyChartGroupOptionsToXml(container, 'barChart', data({}), localName);
		expect(container['c:gapWidth']).toStrictEqual({ '@_val': 'oops' });
		const ofPie: XmlObject = { 'c:gapWidth': { '@_val': '100' } };
		applyChartGroupOptionsToXml(
			ofPie,
			'ofPieChart',
			data({ chartType: 'ofPie', barGapWidth: 5 }),
			localName,
		);
		expect(ofPie['c:gapWidth']).toStrictEqual({ '@_val': '100' });
	});

	it('keeps an out-of-schema authored value PowerPoint wrote when the model matches it', () => {
		const container: XmlObject = { 'c:holeSize': { '@_val': '0' } };
		applyChartGroupOptionsToXml(
			container,
			'doughnutChart',
			data({ chartType: 'doughnut', doughnutHoleSize: 0 }),
			localName,
		);
		expect(container['c:holeSize']).toStrictEqual({ '@_val': '0' });
	});
});

describe('chartGroupOptionVal', () => {
	it('rounds and clamps into the schema range', () => {
		const [gap, overlap, , hole] = CHART_GROUP_OPTIONS;
		expect(chartGroupOptionVal(gap!, 600)).toBe('500');
		expect(chartGroupOptionVal(overlap!, -100.6)).toBe('-100');
		expect(chartGroupOptionVal(hole!, 0)).toBe('1');
	});
});
