/**
 * Parity guard for the chart renderer: the view model of every chart in the committed PowerPoint
 * fixtures, and of a synthetic matrix covering every chart kind, is hashed and compared with a
 * golden file. Moving a helper into `ooxml-core/chart` must leave every hash untouched.
 *
 * Regenerate deliberately with `UPDATE_CHART_PARITY=1 bunx vitest run chart-parity` in `src/ui`.
 */
import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import type { ChartPptxElement, PptxChartData, PptxElement } from 'ooxml-core/pptx';
import { PptxHandler } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildChartViewModel } from './chart-view-model';

const GOLDEN = fileURLToPath(new URL('./__parity__/chart-view-models.json', import.meta.url));
const FIXTURES = fileURLToPath(
	new URL('../../../../core/pptx/__tests__/fixtures/', import.meta.url),
);

const hash = (value: unknown): string =>
	createHash('sha1')
		.update(JSON.stringify(value) ?? 'undefined')
		.digest('hex')
		.slice(0, 16);

function collect(elements: PptxElement[], out: ChartPptxElement[]): void {
	for (const el of elements) {
		if (el.type === 'chart') out.push(el as ChartPptxElement);
		else if (el.type === 'group') collect((el as { children: PptxElement[] }).children, out);
	}
}

const chartElement = (chartData: PptxChartData): ChartPptxElement =>
	({
		type: 'chart',
		id: 'c1',
		x: 0,
		y: 0,
		width: 480,
		height: 320,
		chartData,
	}) as unknown as ChartPptxElement;

const cats = ['Q1', 'Q2', 'Q3', 'Q4'];
const base = (chartType: string, extra: Record<string, unknown> = {}): PptxChartData =>
	({
		chartType,
		title: `${chartType} chart`,
		categories: cats,
		series: [
			{ name: 'North', values: [12, 19, 7, 25] },
			{ name: 'South', values: [8, 14, 22, 11] },
			{ name: 'East', values: [3, 9, 15, 6] },
		],
		style: { hasLegend: true, legendPosition: 'b' },
		...extra,
	}) as unknown as PptxChartData;

function synthetic(): Record<string, PptxChartData> {
	const fmt = (numberFormat: string) => ({
		series: [
			{ name: 'Rate', values: [0.1, 0.25, 0.4, 0.55], numberFormat },
			{ name: 'Base', values: [0.2, 0.3, 0.5, 0.6], numberFormat },
		],
	});
	return {
		bar: base('bar'),
		barStacked: base('bar', { grouping: 'stacked' }),
		barPercent: base('bar', { grouping: 'percentStacked' }),
		barHorizontal: base('bar', { barDirection: 'bar' }),
		barPercentFormat: base('bar', fmt('0%')),
		barCurrencyFormat: base('bar', fmt('$#,##0.00')),
		line: base('line'),
		pie: base('pie'),
		doughnut: base('doughnut'),
		area: base('area'),
		scatter: base('scatter', {
			series: [
				{ name: 'A', values: [3, 5, 4, 8], xValues: [1, 2, 3, 4] },
				{ name: 'B', values: [6, 2, 9, 1], xValues: [1.5, 2.5, 3.5, 4.5] },
			],
		}),
		bubble: base('bubble', {
			series: [
				{ name: 'A', values: [3, 5, 4, 8], xValues: [1, 2, 3, 4], bubbleSizes: [5, 10, 3, 8] },
				{
					name: 'B',
					values: [6, 2, 9, 1],
					xValues: [1.5, 2.5, 3.5, 4.5],
					bubbleSizes: [2, 4, 9, 1],
				},
			],
		}),
		radar: base('radar'),
		stock: base('stock', {
			series: [
				{ name: 'Open', values: [10, 11, 12, 11] },
				{ name: 'High', values: [14, 15, 15, 14] },
				{ name: 'Low', values: [9, 10, 10, 9] },
				{ name: 'Close', values: [12, 13, 11, 13] },
			],
		}),
		stockHlc: base('stock', {
			series: [
				{ name: 'High', values: [14, 15, 15, 14] },
				{ name: 'Low', values: [9, 10, 10, 9] },
				{ name: 'Close', values: [12, 13, 11, 13] },
			],
		}),
		combo: base('combo', {
			series: [
				{ name: 'Sales', values: [12, 19, 7, 25], seriesChartType: 'bar' },
				{ name: 'Trend', values: [8, 14, 22, 11], seriesChartType: 'line', axisId: 2 },
			],
		}),
		bar3D: base('bar3D'),
		line3D: base('line3D'),
		pie3D: base('pie3D'),
		area3D: base('area3D'),
		surface: base('surface'),
		histogram: base('histogram', {
			series: [{ name: 'Data', values: [1, 2, 2, 3, 3, 3, 4, 4, 5, 9, 10, 10] }],
		}),
		waterfall: base('waterfall', { series: [{ name: 'Flow', values: [100, -30, 45, -20] }] }),
		funnel: base('funnel', { series: [{ name: 'Stage', values: [100, 70, 40, 15] }] }),
		treemap: base('treemap', { series: [{ name: 'Size', values: [30, 20, 10, 5] }] }),
		sunburst: base('sunburst', { series: [{ name: 'Size', values: [30, 20, 10, 5] }] }),
		boxWhisker: base('boxWhisker', {
			series: [{ name: 'Data', values: [1, 2, 2, 3, 3, 3, 4, 4, 5, 9] }],
		}),
		ofPie: base('ofPie'),
		regionMap: base('regionMap'),
		unknown: base('unknown'),
		bubbleWidthNegative: base('bubble', {
			bubbleOptions: { sizeRepresents: 'w', bubbleScale: 150, showNegativeBubbles: true },
			series: [
				{ name: 'A', values: [3, 5, 4, 8], xValues: [1, 2, 3, 4], bubbleSizes: [5, -10, 3, 8] },
			],
		}),
		bubbleHiddenNegative: base('bubble', {
			series: [
				{ name: 'A', values: [3, 5, 4, 8], xValues: [1, 2, 3, 4], bubbleSizes: [5, -10, 3, 8] },
			],
		}),
		stockStyled: base('stock', {
			hiLowLines: { color: '#FF0000', width: 2 },
			upDownBars: { upBars: { fillColor: '#00AA00' }, downBars: { fillColor: '#AA0000' } },
			series: [
				{ name: 'Open', values: [10, 11, 12, 11] },
				{ name: 'High', values: [14, 15, 15, 14] },
				{ name: 'Low', values: [9, 10, 10, 9] },
				{ name: 'Close', values: [12, 10, 13, 12] },
			],
		}),
		stockVolume: base('combo', {
			series: [
				{ name: 'Volume', values: [100, 200, 150, 120], seriesChartType: 'bar' },
				{ name: 'Open', values: [10, 11, 12, 11], seriesChartType: 'stock', axisId: 2 },
				{ name: 'High', values: [14, 15, 15, 14], seriesChartType: 'stock', axisId: 2 },
				{ name: 'Low', values: [9, 10, 10, 9], seriesChartType: 'stock', axisId: 2 },
				{ name: 'Close', values: [12, 10, 13, 12], seriesChartType: 'stock', axisId: 2 },
			],
		}),
		treemapLevels: base('treemap', {
			categories: ['a', 'b', 'c', 'd', 'e'],
			categoryLevels: [
				['a', 'b', 'c', 'd', 'e'],
				['Fruit', '', 'Veg', '', 'Nuts'],
			],
			series: [{ name: 'Size', values: [30, 20, 10, 5, 8] }],
		}),
		treemapSeries: base('treemap', {
			series: [
				{ name: 'One', values: [30, 20, 10, 5] },
				{ name: 'Two', values: [12, 8, 6, 4] },
			],
		}),
		sunburstLevels: base('sunburst', {
			categories: ['a', 'b', 'c', 'd', 'e'],
			categoryLevels: [
				['a', 'b', 'c', 'd', 'e'],
				['Fruit', '', 'Veg', '', 'Nuts'],
			],
			series: [{ name: 'Size', values: [30, 20, 10, 5, 8] }],
		}),
		sunburstSeries: base('sunburst', {
			series: [
				{ name: 'One', values: [30, 20, 10, 5] },
				{ name: 'Two', values: [12, 8, 6, 4] },
			],
		}),
		legendRight: base('bar', { style: { hasLegend: true, legendPosition: 'r' } }),
		legendTop: base('line', { style: { hasLegend: true, legendPosition: 't' } }),
		noLegend: base('bar', { style: { hasLegend: false } }),
	};
}

async function snapshot(): Promise<Record<string, string>> {
	const out: Record<string, string> = {};
	for (const [name, data] of Object.entries(synthetic()))
		out[`synthetic/${name}`] = hash(buildChartViewModel(chartElement(data)));
	const files = readdirSync(FIXTURES, { recursive: true, encoding: 'utf8' })
		.map((f) => f.split('\\').join('/'))
		.filter((f) => f.endsWith('.pptx'))
		.sort();
	for (const file of files) {
		const bytes = readFileSync(`${FIXTURES}${file}`);
		let data;
		try {
			data = await new PptxHandler().load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
			);
		} catch {
			continue;
		}
		if (!data) continue;
		data.slides.forEach((slide, s) => {
			const charts: ChartPptxElement[] = [];
			collect(slide.elements, charts);
			charts.forEach((chart, i) => {
				out[`${file}#${s + 1}.${i + 1}`] = hash(buildChartViewModel(chart));
			});
		});
	}
	return out;
}

describe('chart view-model parity', () => {
	it('matches the committed golden hashes', async () => {
		const actual = await snapshot();
		if (process.env['UPDATE_CHART_PARITY'] || !existsSync(GOLDEN)) {
			writeFileSync(GOLDEN, `${JSON.stringify(actual, null, '\t')}\n`);
		}
		const golden = JSON.parse(readFileSync(GOLDEN, 'utf8')) as Record<string, string>;
		expect(actual).toEqual(golden);
		expect(Object.keys(actual).filter((k) => k.startsWith('synthetic/')).length).toBeGreaterThan(
			30,
		);
		expect(Object.keys(actual).filter((k) => !k.startsWith('synthetic/')).length).toBeGreaterThan(
			10,
		);
	}, 120_000);
});
