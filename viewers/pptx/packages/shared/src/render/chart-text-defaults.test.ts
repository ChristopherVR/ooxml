import type { PptxChartData } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { chartTextFontFamily } from './chart-font';
import { resolveChartTextStyle, withChartTextDefaults } from './chart-text-defaults';
import { buildChartViewModel } from './chart-view-model-build';

function chart(overrides: Partial<PptxChartData> = {}): PptxChartData {
	return {
		chartType: 'bar',
		categories: ['A'],
		series: [{ name: 'S', values: [1] }],
		...overrides,
	};
}

const chartWide = { fontSize: 8, fontFamily: 'Arial', eastAsiaFontFamily: 'Malgun Gothic' };

describe('resolveChartTextStyle', () => {
	it('lets each level override only the fields it sets', () => {
		const data = chart({ style: { textStyle: chartWide } });
		expect(
			resolveChartTextStyle(data, { fontSize: 12, color: '#112233' }, { bold: true }),
		).toStrictEqual({ ...chartWide, fontSize: 12, color: '#112233', bold: true });
	});

	it('keeps the chart-wide Latin face for a level that names only an East Asian one', () => {
		const data = chart({ style: { textStyle: { fontFamily: 'Calibri' } } });
		expect(resolveChartTextStyle(data, { eastAsiaFontFamily: 'Batang' })).toStrictEqual({
			fontFamily: 'Calibri',
			eastAsiaFontFamily: 'Batang',
		});
	});

	it('returns undefined when no level sets anything', () => {
		expect(resolveChartTextStyle(chart(), undefined, {})).toBeUndefined();
	});
});

describe('withChartTextDefaults', () => {
	it('gives an axis without its own font the chart-wide style', () => {
		const data = chart({
			style: { textStyle: chartWide },
			axes: [{ axisType: 'catAx' }, { axisType: 'valAx', fontSize: 12, fontFamily: 'Georgia' }],
		});
		const axes = withChartTextDefaults(data).axes;
		expect(axes?.[0]).toMatchObject({ fontSize: 8, fontFamily: chartTextFontFamily(chartWide) });
		expect(axes?.[1]).toMatchObject({
			fontSize: 12,
			fontFamily: chartTextFontFamily({ ...chartWide, fontFamily: 'Georgia' }),
		});
	});

	it('returns the same object when there is nothing to resolve', () => {
		const data = chart({ axes: [{ axisType: 'catAx' }] });
		expect(withChartTextDefaults(data)).toBe(data);
	});

	it('does not touch the model the editor saves', () => {
		const data = chart({ style: { textStyle: { fontSize: 8 } }, axes: [{ axisType: 'catAx' }] });
		withChartTextDefaults(data);
		expect(data.axes?.[0].fontSize).toBeUndefined();
	});
});

describe('buildChartViewModel with a chart-wide text style', () => {
	const element = (chartData: PptxChartData) => ({
		id: 'c',
		type: 'chart' as const,
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		chartData,
	});

	it('draws axis labels, data labels and legend entries with it', () => {
		const vm = buildChartViewModel(
			element(
				chart({
					categories: ['서울', '부산'],
					series: [{ name: '매출', values: [3, 2] }],
					style: {
						textStyle: chartWide,
						hasLegend: true,
						hasDataLabels: true,
						dataLabels: { showValue: true },
					},
					axes: [{ axisType: 'catAx' }, { axisType: 'valAx' }],
				}),
			),
		);
		const family = chartTextFontFamily(chartWide);
		expect(vm.axisLabels.some((label) => label.fontFamily === family)).toBeTruthy();
		expect(vm.dataLabels.every((label) => label.fontFamily === family)).toBeTruthy();
		expect(
			vm.legend.every((entry) => entry.textStyle?.eastAsiaFontFamily === 'Malgun Gothic'),
		).toBeTruthy();
	});

	it('draws data table cells and the title with it', () => {
		const vm = buildChartViewModel(
			element(
				chart({
					title: '매출',
					style: { textStyle: chartWide, hasTitle: true, titleEastAsiaFontFamily: 'Batang' },
					dataTable: { showKeys: true },
				}),
			),
		);
		expect(vm.titleStyle?.fontFamily).toBe(
			chartTextFontFamily({ fontFamily: 'Arial', eastAsiaFontFamily: 'Batang' }),
		);
		const cells = vm.dataTable?.filter((p) => p.kind === 'text') ?? [];
		expect(cells.length).toBeGreaterThan(0);
		expect(cells.every((cell) => cell.fontFamily === chartTextFontFamily(chartWide))).toBeTruthy();
		// 8pt chart-wide size, in px.
		expect(cells.every((cell) => cell.fontSize === (8 * 4) / 3)).toBeTruthy();
	});
});
