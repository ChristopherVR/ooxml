import { describe, expect, it } from 'vitest';

import type { PptxChartData, PptxChartType, XmlObject } from '../types';
import { AXIS_LOCAL_NAMES } from './chart-container-content-model';
import { buildChartExSpaceXml, canGenerateChartEx } from './chart-cx-generator';
import { buildChartSpaceXml } from './chart-xml-generator';

// ECMA-376 dml-chart.xsd: line3D/surface3D require 3; bar3D/area3D/surface
// allow 2..3. Other cartesian containers require 2; pie containers forbid axes.
const CASES: Array<[PptxChartType, string, number, number, number]> = [
	['bar', 'barChart', 2, 2, 2],
	['bar3D', 'bar3DChart', 2, 3, 2],
	['line', 'lineChart', 2, 2, 2],
	['line3D', 'line3DChart', 3, 3, 3],
	['pie', 'pieChart', 0, 0, 0],
	['pie3D', 'pie3DChart', 0, 0, 0],
	['ofPie', 'ofPieChart', 0, 0, 0],
	['doughnut', 'doughnutChart', 0, 0, 0],
	['area', 'areaChart', 2, 2, 2],
	['area3D', 'area3DChart', 2, 3, 2],
	['scatter', 'scatterChart', 2, 2, 2],
	['bubble', 'bubbleChart', 2, 2, 2],
	['radar', 'radarChart', 2, 2, 2],
	['stock', 'stockChart', 2, 2, 2],
	['surface', 'surfaceChart', 2, 3, 3],
	['combo', 'barChart', 2, 2, 2],
	['unknown', 'barChart', 2, 2, 2],
];

function data(chartType: PptxChartType): PptxChartData {
	return {
		chartType,
		categories: ['A', 'B', 'C'],
		series: [{ name: 'Series A', values: [0, 2, 3] }],
	};
}

function plotArea(chartData: PptxChartData): XmlObject {
	const root = buildChartSpaceXml(chartData)['c:chartSpace'] as XmlObject;
	return (root['c:chart'] as XmlObject)['c:plotArea'] as XmlObject;
}

function nodes(value: unknown): XmlObject[] {
	return value === undefined ? [] : Array.isArray(value) ? value : [value as XmlObject];
}

function assertAxisSet(plot: XmlObject, tag: string, count: number): void {
	const references = nodes((plot[`c:${tag}`] as XmlObject)['c:axId']).map((n) => n['@_val']);
	const axes = AXIS_LOCAL_NAMES.flatMap((name) => nodes(plot[`c:${name}`]));
	const ids = axes.map((axis) => (axis['c:axId'] as XmlObject)['@_val']);
	expect(references).toHaveLength(count);
	expect(new Set(references).size).toBe(count);
	expect(axes).toHaveLength(count);
	for (const id of references) {
		expect(ids.filter((axisId) => axisId === id)).toHaveLength(1);
	}
	for (const axis of axes) {
		const id = (axis['c:axId'] as XmlObject)['@_val'];
		const cross = (axis['c:crossAx'] as XmlObject)['@_val'];
		expect(cross).not.toBe(id);
		expect(ids.filter((axisId) => axisId === cross)).toHaveLength(1);
		expect((axis['c:scaling'] as XmlObject)['c:orientation']).toBeDefined();
		expect(axis['c:delete']).toBeDefined();
		expect(axis['c:axPos']).toBeDefined();
	}
	if (count === 3) {
		const category = plot['c:catAx'] as XmlObject;
		const value = plot['c:valAx'] as XmlObject;
		const series = plot['c:serAx'] as XmlObject;
		expect(category['c:crossAx']).toEqual(value['c:axId']);
		expect(value['c:crossAx']).toEqual(category['c:axId']);
		expect(series['c:crossAx']).toEqual(value['c:axId']);
	}
}

describe('generated chart axis cardinality and references', () => {
	it.each(CASES)('%s has a complete schema-valid axis set', (type, tag, min, max, count) => {
		const plot = plotArea(data(type));
		assertAxisSet(plot, tag, count);
		expect(count).toBeGreaterThanOrEqual(min);
		expect(count).toBeLessThanOrEqual(max);
		const series = nodes((plot[`c:${tag}`] as XmlObject)['c:ser'])[0];
		const values = (series['c:val'] ?? series['c:yVal']) as XmlObject;
		expect(nodes((values['c:numLit'] as XmlObject)['c:pt'])[0]['c:v']).toBe('0');
	});

	it('generates the 3-D surface projection with three complete axes', () => {
		assertAxisSet(plotArea({ ...data('surface'), surfaceTopView: false }), 'surface3DChart', 3);
	});

	it('applies series-axis formatting without category/value-only children', () => {
		const plot = plotArea({
			...data('line3D'),
			axes: [
				{
					axisType: 'serAx',
					deleted: true,
					orientation: 'maxMin',
					axPos: 't',
					tickLabelSkip: 2,
					tickMarkSkip: 3,
				},
			],
		});
		const axis = plot['c:serAx'] as XmlObject;
		expect(axis['c:delete']).toEqual({ '@_val': '1' });
		expect((axis['c:scaling'] as XmlObject)['c:orientation']).toEqual({ '@_val': 'maxMin' });
		expect(axis['c:axPos']).toEqual({ '@_val': 't' });
		expect(axis['c:tickLblSkip']).toEqual({ '@_val': '2' });
		expect(axis['c:tickMarkSkip']).toEqual({ '@_val': '3' });
		expect(axis['c:crossBetween']).toBeUndefined();
		expect(axis['c:lblAlgn']).toBeUndefined();
	});

	it.each([
		'histogram',
		'waterfall',
		'funnel',
		'treemap',
		'sunburst',
		'boxWhisker',
		'regionMap',
	] as const)('%s uses ChartEx, not classic axis references', (type) => {
		const chartData = data(type);
		expect(canGenerateChartEx(chartData)).toBe(true);
		const root = buildChartExSpaceXml(chartData)['cx:chartSpace'] as XmlObject;
		const plot = (root['cx:chart'] as XmlObject)['cx:plotArea'] as XmlObject;
		expect(plot['c:axId']).toBeUndefined();
		for (const name of AXIS_LOCAL_NAMES) {
			expect(plot[`c:${name}`]).toBeUndefined();
		}
	});
});
