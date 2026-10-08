import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-transparency.json';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { saveXlsx } from '../write/index';
import { loadXlsx } from '../read/index';
import { chartView } from '../layout/chart-view';
import { renderChartSvg } from '../../chart/render/chart-svg';
import type { ChartObject } from '../model';
import { chartSeriesTransparency, chartSeriesTransparencyPatch } from './chart-series-transparency';
import { chartSeriesFillPatch, chartSeriesSolidFillPatch } from './chart-series-fill';

for (const sample of native.cases)
	it(`renders and round-trips native ${sample.percent}% transparency`, async () => {
		const book = createWorkbook();
		createEditSession(book).addChart(0, {
			chartType: 'column',
			series: [],
			showLegend: false,
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		});
		const zip = await JSZip.loadAsync(await saveXlsx(book));
		zip.file('xl/charts/chart1.xml', sample.chartXml);
		const loaded = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const chart = loaded.sheets[0]!.drawings[0] as ChartObject;
		expect(sample.transparency).toBeCloseTo(sample.percent / 100, 6);
		expect(chartSeriesTransparency(loaded.sheets[0]!.drawings[0] as ChartObject, 0)).toBe(
			sample.percent,
		);
		const color = sample.percent === 0 ? '#FF0000' : `rgba(255,0,0,${1 - sample.percent / 100})`;
		const view = chartView(loaded, 0, chart, () => []);
		expect(view.series[0]!.color).toBe(color);
		expect(renderChartSvg(view, 480, 300)).toContain(`fill="${color}"`);
		const session = createEditSession(loaded);
		session.updateChart(0, 0, chartSeriesTransparencyPatch(chart, 0, 23)!);
		expect(chartSeriesTransparency(chart, 0)).toBe(23);
		const saved = await loadXlsx(await saveXlsx(loaded));
		expect(chartSeriesTransparency(saved.sheets[0]!.drawings[0] as ChartObject, 0)).toBe(23);
		session.undo();
		expect(chartSeriesTransparency(loaded.sheets[0]!.drawings[0] as ChartObject, 0)).toBe(
			sample.percent,
		);
	});

it('preserves opacity on color edits, replaces ordered alpha transforms and guards unsupported fills', () => {
	const chart: ChartObject = {
		kind: 'chart',
		chartType: 'column',
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [
			{
				categories: ['A'],
				values: [1],
				drawingColor: {
					kind: 'scheme',
					value: 'accent1',
					transforms: [
						{ name: 'lumMod', value: '60000' },
						{ name: 'lumOff', value: '40000' },
						{ name: 'alpha', value: '50000' },
						{ name: 'alphaMod', value: '80000' },
						{ name: 'alphaOff', value: '10000' },
					],
				},
			},
		],
	};
	expect(chartSeriesTransparency(chart, 0)).toBe(50);
	const patch = chartSeriesTransparencyPatch(chart, 0, 37)!;
	expect(patch.series![0]!.drawingColor!.transforms).toEqual([
		{ name: 'lumMod', value: '60000' },
		{ name: 'lumOff', value: '40000' },
		{ name: 'alpha', value: '63000' },
	]);
	const edited = { ...chart, ...patch };
	expect(chartSeriesTransparencyPatch(edited, 0, 37)).toBeUndefined();
	const recolored = chartSeriesFillPatch(edited, 0, { rgb: '0000FF' })!;
	expect(chartSeriesTransparency({ ...edited, ...recolored }, 0)).toBeCloseTo(
		native.colorChangeTransparency * 100,
		5,
	);
	for (const percent of [-1, 101, NaN])
		expect(() => chartSeriesTransparencyPatch(chart, 0, percent)).toThrow(RangeError);
	expect(() => chartSeriesTransparencyPatch(chart, -1, 0)).toThrow(RangeError);
	chart.series[0]!.fill = { kind: 'none' };
	expect(chartSeriesTransparency(chart, 0)).toBeUndefined();
	expect(chartSeriesTransparencyPatch(chart, 0, 37)).toBeUndefined();
	const color = {
		kind: 'srgb' as const,
		value: 'FF0000',
		transforms: [{ name: 'alpha', value: '63000' }],
	};
	chart.series[0]!.fill = {
		kind: 'gradient',
		stops: [
			{ position: 0, color },
			{ position: 1, color },
		],
	};
	expect(chartSeriesTransparency(chart, 0)).toBeUndefined();
	const solid = chartSeriesSolidFillPatch(chart, 0)!;
	expect(solid.series![0]!.drawingColor).toEqual(color);
	expect(solid.series![0]!.drawingColor).not.toBe(color);
	chart.series[0]!.fill = { kind: 'solid', color };
	chart.series[0]!.pointFills = { 0: { kind: 'solid', color } };
	expect(chartSeriesTransparency(chart, 0)).toBe(37);
	expect(chartView(createWorkbook(), 0, chart, () => []).series[0]!.pointColors![0]).toBe(
		'rgba(255,0,0,0.63)',
	);
});
