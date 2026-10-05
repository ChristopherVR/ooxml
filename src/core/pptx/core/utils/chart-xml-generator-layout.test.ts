import { describe, expect, it } from 'vitest';

import type { PptxChartData, XmlObject } from '../types';
import { buildChartSpaceXml } from './chart-xml-generator';

function makeData(overrides?: Partial<PptxChartData>): PptxChartData {
	return {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [{ name: 'Revenue', values: [1, 2] }],
		...overrides,
	};
}

function chartSpace(data: PptxChartData): XmlObject {
	return buildChartSpaceXml(data)['c:chartSpace'] as XmlObject;
}

function plotArea(data: PptxChartData): XmlObject {
	return (chartSpace(data)['c:chart'] as XmlObject)['c:plotArea'] as XmlObject;
}

function container(data: PptxChartData, tag: string): XmlObject {
	return plotArea(data)[tag] as XmlObject;
}

function keys(node: XmlObject): string[] {
	return Object.keys(node).filter((key) => !key.startsWith('@_'));
}

describe('buildChartSpaceXml: unset layout fields keep the previous output', () => {
	it('writes gapWidth 150, no overlap, delete 0 and no chart-area spPr or roundedCorners', () => {
		const space = chartSpace(makeData());
		const bar = container(makeData(), 'c:barChart');
		expect(bar['c:gapWidth']).toStrictEqual({ '@_val': '150' });
		expect(bar['c:overlap']).toBeUndefined();
		expect(keys(bar)).toStrictEqual([
			'c:barDir',
			'c:grouping',
			'c:varyColors',
			'c:ser',
			'c:gapWidth',
			'c:axId',
		]);
		const area = plotArea(makeData());
		expect((area['c:catAx'] as XmlObject)['c:delete']).toStrictEqual({ '@_val': '0' });
		expect((area['c:valAx'] as XmlObject)['c:delete']).toStrictEqual({ '@_val': '0' });
		expect(area['c:spPr']).toBeUndefined();
		expect(space['c:spPr']).toBeUndefined();
		expect(space['c:roundedCorners']).toBeUndefined();
	});

	it('writes holeSize 50 and no firstSliceAng for a doughnut', () => {
		const doughnut = container(makeData({ chartType: 'doughnut' }), 'c:doughnutChart');
		expect(doughnut['c:holeSize']).toStrictEqual({ '@_val': '50' });
		expect(doughnut['c:firstSliceAng']).toBeUndefined();
	});
});

describe('buildChartSpaceXml: model layout fields', () => {
	it('writes gapWidth and overlap from the model in CT_BarChart order', () => {
		const bar = container(makeData({ barGapWidth: 35, barOverlap: -4 }), 'c:barChart');
		expect(bar['c:gapWidth']).toStrictEqual({ '@_val': '35' });
		expect(bar['c:overlap']).toStrictEqual({ '@_val': '-4' });
		expect(keys(bar)).toStrictEqual([
			'c:barDir',
			'c:grouping',
			'c:varyColors',
			'c:ser',
			'c:gapWidth',
			'c:overlap',
			'c:axId',
		]);
	});

	it('writes gapWidth but no overlap for bar3D (CT_Bar3DChart has none)', () => {
		const bar = container(
			makeData({ chartType: 'bar3D', barGapWidth: 80, barOverlap: 20, barShape: 'cone' }),
			'c:bar3DChart',
		);
		expect(bar['c:gapWidth']).toStrictEqual({ '@_val': '80' });
		expect(bar['c:overlap']).toBeUndefined();
		expect(keys(bar).slice(-3)).toStrictEqual(['c:gapWidth', 'c:shape', 'c:axId']);
	});

	it('writes firstSliceAng and holeSize in CT_DoughnutChart order', () => {
		const doughnut = container(
			makeData({ chartType: 'doughnut', firstSliceAngle: 90, doughnutHoleSize: 20 }),
			'c:doughnutChart',
		);
		expect(keys(doughnut)).toStrictEqual([
			'c:varyColors',
			'c:ser',
			'c:firstSliceAng',
			'c:holeSize',
		]);
		expect(doughnut['c:firstSliceAng']).toStrictEqual({ '@_val': '90' });
		expect(doughnut['c:holeSize']).toStrictEqual({ '@_val': '20' });
	});

	it('writes firstSliceAng for a pie and the of-pie gap width', () => {
		const pie = container(makeData({ chartType: 'pie', firstSliceAngle: 270 }), 'c:pieChart');
		expect(pie['c:firstSliceAng']).toStrictEqual({ '@_val': '270' });
		const ofPie = container(
			makeData({ chartType: 'ofPie', ofPieOptions: { ofPieType: 'bar', gapWidth: 60 } }),
			'c:ofPieChart',
		);
		expect(ofPie['c:gapWidth']).toStrictEqual({ '@_val': '60' });
	});

	it('rounds and clamps out-of-schema model values instead of writing invalid XML', () => {
		const doughnut = container(
			makeData({ chartType: 'doughnut', doughnutHoleSize: 0, firstSliceAngle: 400.4 }),
			'c:doughnutChart',
		);
		expect(doughnut['c:holeSize']).toStrictEqual({ '@_val': '1' });
		expect(doughnut['c:firstSliceAng']).toStrictEqual({ '@_val': '360' });
	});

	it('writes c:delete from axes[].deleted', () => {
		const area = plotArea(
			makeData({
				axes: [
					{ axisType: 'catAx', deleted: true },
					{ axisType: 'valAx', deleted: false },
				],
			}),
		);
		const catAx = area['c:catAx'] as XmlObject;
		expect(catAx['c:delete']).toStrictEqual({ '@_val': '1' });
		expect(keys(catAx).slice(0, 4)).toStrictEqual(['c:axId', 'c:scaling', 'c:delete', 'c:axPos']);
		expect((area['c:valAx'] as XmlObject)['c:delete']).toStrictEqual({ '@_val': '0' });
	});

	it('writes the chart-area spPr after c:chart and roundedCorners before it', () => {
		const space = chartSpace(
			makeData({
				roundedCorners: false,
				style: { chartAreaFill: 'none', chartAreaBorder: 'none' },
				protection: { chartObject: true },
			}),
		);
		expect(keys(space)).toStrictEqual(['c:roundedCorners', 'c:protection', 'c:chart', 'c:spPr']);
		expect(space['c:roundedCorners']).toStrictEqual({ '@_val': '0' });
		expect(space['c:spPr']).toStrictEqual({ 'a:noFill': {}, 'a:ln': { 'a:noFill': {} } });
	});

	it('writes solid and gradient area fills and a coloured border', () => {
		const solid = chartSpace(
			makeData({ style: { chartAreaFill: '#ffffff', chartAreaBorder: '#000000' } }),
		);
		expect(solid['c:spPr']).toStrictEqual({
			'a:solidFill': { 'a:srgbClr': { '@_val': 'FFFFFF' } },
			'a:ln': { 'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } } },
		});
		const gradient = chartSpace(
			makeData({
				style: {
					chartAreaFill: '#FF0000',
					chartAreaGradient: {
						type: 'linear',
						angle: 90,
						stops: [
							{ color: '#FFFFFF', position: 0 },
							{ color: '#000000', position: 100 },
						],
					},
				},
			}),
		);
		expect(Object.keys(gradient['c:spPr'] as XmlObject)).toStrictEqual(['a:gradFill']);
	});

	it('writes the plot-area spPr after the axes', () => {
		const area = plotArea(makeData({ style: { plotAreaFill: 'none', plotAreaBorder: 'none' } }));
		expect(keys(area).slice(-1)).toStrictEqual(['c:spPr']);
		expect(area['c:spPr']).toStrictEqual({ 'a:noFill': {}, 'a:ln': { 'a:noFill': {} } });
	});
});
