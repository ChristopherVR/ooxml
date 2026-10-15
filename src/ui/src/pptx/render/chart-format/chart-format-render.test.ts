import type { ChartPptxElement, PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildChartViewModel } from '../chart-view-model';
import type { ChartViewModel, SvgPrimitive } from '../chart-view-model';
import { columnChart } from './chart-format-fixtures.test-util';

function element(chartData: PptxChartData): ChartPptxElement {
	return {
		id: 'c1',
		type: 'chart',
		x: 0,
		y: 0,
		width: 480,
		height: 300,
		chartData,
	} as ChartPptxElement;
}

function chart(formats: PptxChartData['elementFormats'], extra: Partial<PptxChartData> = {}) {
	const data = { ...columnChart(), ...extra, elementFormats: formats };
	data.axes = data.axes!.slice(0, 2);
	return buildChartViewModel(element(data));
}

const image = (vm: ChartViewModel, suffix: string) => {
	const def = vm.defs?.find((d) => d.id.endsWith(`-fmt-${suffix}`));
	return def && 'href' in def ? decodeURIComponent(def.href) : '';
};
const marks = (vm: ChartViewModel, series: number) =>
	vm.primitives.filter(
		(p): p is Extract<SvgPrimitive, { kind: 'rect' }> =>
			p.kind === 'rect' && p.part?.seriesIndex === series,
	);

describe('chart element formats on the slide', () => {
	it('leaves a chart whose formats draw nothing new unchanged', () => {
		const plain = chart(undefined);
		const none = chart({
			plotArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
			legend: { fill: { kind: 'none' } },
			'axis:1': { fill: { kind: 'none' }, line: { kind: 'solid', color: '#BFBFBF' } },
		});
		expect(none).toEqual(plain);
	});

	it('paints transparent series marks', () => {
		const vm = chart({
			'series:0': { fill: { kind: 'solid', color: '#4472C4', transparency: 40 } },
		});
		expect(marks(vm, 0).every((m) => m.opacity === 0.6)).toBe(true);
		expect(marks(vm, 1).every((m) => (m.opacity ?? 1) === 1)).toBe(true);
	});

	it('casts a blurred, offset shadow and a glow behind a series', () => {
		const vm = chart({
			'series:1': {
				effects: {
					shadow: { color: '#000000', transparency: 60, blur: 6, distance: 3, angle: 45 },
					glow: { color: '#FFC000', radius: 5 },
				},
			},
		});
		const svg = image(vm, 's1');
		expect(svg).toContain('<feGaussianBlur stdDeviation="4"/>');
		expect(svg).toMatch(/opacity="0.4" transform="translate\(2.83 2.83\)"/u);
		expect(svg).toContain('stroke="#FFC000"');
		const at = vm.primitives.findIndex((p) => p.kind === 'rect' && p.fill === 'url(#c1-fmt-s1)');
		const firstMark = vm.primitives.findIndex((p) => 'part' in p && p.part?.seriesIndex === 1);
		expect(at).toBeGreaterThanOrEqual(0);
		expect(at).toBe(firstMark - 1);
	});

	it('fades soft-edged bars and keeps the originals for hit-testing', () => {
		const vm = chart({ 'series:0': { effects: { softEdge: 4 } } });
		expect(image(vm, 's0')).toContain('<mask');
		expect(marks(vm, 0).every((m) => m.opacity === 0)).toBe(true);
	});

	it('draws the plot area, legend, title and axis label boxes in the chart background', () => {
		const vm = chart(
			{
				plotArea: {
					fill: { kind: 'solid', color: '#F2F2F2', transparency: 20 },
					line: { kind: 'solid', color: '#7F7F7F' },
				},
				legend: {
					fill: { kind: 'solid', color: '#FFF2CC' },
					effects: { shadow: { color: '#000000', blur: 4, distance: 3, angle: 90 } },
				},
				title: { line: { kind: 'solid', color: '#000000', width: 1 } },
				'axis:0': { fill: { kind: 'solid', color: '#DDEBF7' } },
			},
			{ style: { ...columnChart().style, chartAreaFill: '#FFFFFF' } },
		);
		expect(vm.areaFill).toBe('url(#c1-fmt-bg)');
		const svg = image(vm, 'bg');
		expect(svg).toContain('fill="#FFFFFF"');
		expect(svg).toContain('fill="#F2F2F2" fill-opacity="0.8"');
		expect(svg).toContain('stroke="#7F7F7F"');
		expect(svg).toContain('fill="#FFF2CC"');
		expect(svg).toContain('fill="#DDEBF7"');
		expect(svg).toContain('stroke="#000000"');
	});

	it('fills data label boxes with the data-label format, transparency included', () => {
		const vm = chart(
			{ dataLabels: { fill: { kind: 'solid', color: '#FF0000', transparency: 50 } } },
			{
				style: { ...columnChart().style, hasDataLabels: true, dataLabels: { showValue: true } },
			},
		);
		const boxes = vm.primitives.filter((p) => p.kind === 'polygon');
		expect(boxes.length).toBeGreaterThan(0);
		expect(boxes.every((p) => p.kind === 'polygon' && p.fill === 'rgba(255, 0, 0, 0.5)')).toBe(
			true,
		);
	});

	it('builds the same view model every time', () => {
		const formats: PptxChartData['elementFormats'] = {
			'series:0': { effects: { glow: { color: '#00B050', radius: 3 } } },
			legend: { fill: { kind: 'solid', color: '#FFF2CC' } },
		};
		expect(chart(formats)).toEqual(chart(formats));
	});
});
