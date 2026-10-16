import type { ChartPptxElement, PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildChartViewModel } from '../chart-view-model';
import type { ChartViewModel, SvgPrimitive } from '../chart-view-model';
import { getComputedEffectStyle } from '../visual-effects';
import { chartAreaEffectStyle } from './chart-area-outer-effects';
import { columnChart } from './chart-format-fixtures.test-util';

const PNG = 'data:image/png;base64,iVBORw0KGgo=';

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
	data.axes = data.axes!.slice(0, 2).map((axis) => ({ ...axis, spPr: { strokeColor: '#808080' } }));
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
const lines = (vm: ChartViewModel) => [...vm.gridlines, ...vm.primitives];

describe('chart element fills, lines, 3-D and text on the slide', () => {
	it('fills a series with a preset pattern, legend swatch included', () => {
		const vm = chart({
			'series:0': {
				fill: {
					kind: 'pattern',
					pattern: { preset: 'ltUpDiag', foreground: '#C00000', background: '#FFFFFF' },
				},
			},
		});
		const fill = marks(vm, 0)[0]!.fill;
		expect(fill).toBe('url(#c1-fmt-pt0)');
		expect(vm.legend[0]!.color).toBe(fill);
		expect(image(vm, 'pt0')).toContain('#C00000');
		expect(marks(vm, 1)[0]!.fill).toBe('#ED7D31');
	});

	it('paints picture, pattern and gradient boxes', () => {
		const vm = chart({
			plotArea: { fill: { kind: 'picture', picture: { src: PNG } } },
			legend: {
				fill: {
					kind: 'gradient',
					gradient: {
						type: 'linear',
						angle: 0,
						stops: [
							{ color: '#FFFFFF', position: 0 },
							{ color: '#5B9BD5', position: 100 },
						],
					},
				},
			},
			title: {
				fill: {
					kind: 'pattern',
					pattern: { preset: 'smGrid', foreground: '#000000', background: '#FFFFFF' },
				},
			},
		});
		const bg = image(vm, 'bg');
		expect(bg).toContain(`<image href="${PNG}"`);
		expect(bg).toContain('<linearGradient');
		expect(bg).toMatch(/<pattern id="pp\d+" patternUnits="userSpaceOnUse"/u);
	});

	it('draws axis arrows and a compound gradient axis line, hiding the plain line', () => {
		const vm = chart({
			'axis:0': {
				line: {
					kind: 'gradient',
					width: 2,
					compound: 'dbl',
					endArrow: 'triangle',
					gradient: {
						type: 'linear',
						angle: 0,
						stops: [
							{ color: '#000000', position: 0 },
							{ color: '#FF0000', position: 100 },
						],
					},
				},
			},
		});
		const ax = image(vm, 'ax');
		expect(ax).toContain('<polygon');
		expect((ax.match(/<line /gu) ?? []).length).toBe(2);
		expect(ax).toContain('<linearGradient');
		expect(lines(vm).some((l) => l.kind === 'line' && l.opacity === 0)).toBe(true);
		// The build-time tags never reach the bindings.
		expect(JSON.stringify(vm)).not.toContain('axisLine');
	});

	it('never leaves an axis-line tag on a chart without formats', () => {
		expect(JSON.stringify(chart(undefined))).not.toContain('axisLine');
	});

	it('shades a beveled series and outlines its contour', () => {
		const vm = chart({
			'series:1': {
				threeD: {
					bevelTop: { type: 'circle', width: 6, height: 6 },
					contour: { size: 1, color: '#FFFFFF' },
					material: 'metal',
				},
			},
		});
		const bevel = image(vm, 'b1');
		expect(bevel).toContain('stop-color="#FFFFFF"');
		expect(bevel).toContain('stop-color="#000000"');
		expect(bevel).toContain('stroke="#FFFFFF"');
		const index = vm.primitives.findIndex((p) => p.kind === 'rect' && p.fill === 'url(#c1-fmt-b1)');
		const last = vm.primitives.findLastIndex((p) => 'part' in p && p.part?.seriesIndex === 1);
		expect(index).toBe(last + 1);
	});

	it('draws text gradient, outline, glow and a custom label angle', () => {
		const vm = chart({
			'axis:0': {
				text: {
					fill: {
						kind: 'gradient',
						gradient: {
							type: 'linear',
							angle: 0,
							stops: [
								{ color: '#FF0000', position: 0 },
								{ color: '#0000FF', position: 100 },
							],
						},
					},
					outline: { kind: 'solid', color: '#000000', width: 0.5 },
					effects: { glow: { color: '#FFC000', radius: 4 } },
					rotation: -45,
				},
			},
		});
		expect(vm.categoryLabels.every((l) => l.opacity === 0)).toBe(true);
		expect(vm.categoryLabels[0]!.transform).toMatch(/^rotate\(-45, /u);
		expect(vm.categoryLabels[0]!.textAnchor).toBe('end');
		const tx = image(vm, 'tx');
		expect(tx).toContain('<linearGradient');
		expect(tx).toContain('stroke="#000000"');
		expect(tx).toContain('feGaussianBlur');
		expect(tx).toContain('>Q1</text>');
	});

	it('casts data label shadows from the label boxes', () => {
		const vm = chart({
			dataLabels: {
				fill: { kind: 'solid', color: '#FFFFFF' },
				effects: { shadow: { color: '#000000', blur: 4, distance: 3, angle: 45 } },
			},
		});
		expect(image(vm, 'tx')).toMatch(/<g filter="url\(#s\d+\)"[^>]*><rect /u);
	});

	it('casts the chart area shadow and glow outside the frame', () => {
		const data: PptxChartData = {
			...columnChart(),
			style: { ...columnChart().style, chartAreaFill: '#FFFFFF' },
			elementFormats: {
				chartArea: {
					effects: {
						shadow: { color: '#000000', transparency: 60, blur: 4, distance: 3, angle: 45 },
						glow: { color: '#FFC000', radius: 5 },
					},
				},
			},
		};
		const style = chartAreaEffectStyle(data)!;
		expect(style.shadowDistance).toBeCloseTo(4);
		expect(style.glowRadius).toBeCloseTo(20 / 3);
		const fx = getComputedEffectStyle(element(data));
		expect(fx.boxShadow).toContain('rgba(0, 0, 0, 0.4)');
		expect(fx.filter).toContain('drop-shadow');
		const noFill = { ...data, style: { ...data.style, chartAreaFill: 'none' } };
		expect(chartAreaEffectStyle(noFill)).toBeUndefined();
	});

	it('clusters a secondary-axis bar series as its own group, drawn on top', () => {
		const base = columnChart();
		const data: PptxChartData = {
			...base,
			series: [
				{ ...base.series[0]!, axisId: 2 },
				{ ...base.series[1]!, axisId: 3 },
			],
			axes: [
				{ axisType: 'catAx', axisId: 1, axPos: 'b', crossAxisId: 2 },
				{ axisType: 'valAx', axisId: 2, axPos: 'l', crossAxisId: 1 },
				{ axisType: 'valAx', axisId: 3, axPos: 'r', crossAxisId: 4, crosses: 'max' },
				{ axisType: 'catAx', axisId: 4, axPos: 'b', crossAxisId: 3, deleted: true },
			],
		};
		const vm = buildChartViewModel(element(data));
		const primary = marks(vm, 0)[0]!;
		const secondary = marks(vm, 1)[0]!;
		// Each group alone in its slot: both span the same cluster, the secondary painted later.
		expect(secondary.w).toBeCloseTo(primary.w);
		expect(secondary.x).toBeCloseTo(primary.x);
		expect(vm.primitives.indexOf(secondary)).toBeGreaterThan(vm.primitives.indexOf(primary));
	});
});
