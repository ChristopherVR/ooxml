// Bubble, stock, treemap and sunburst charts authored by Excel (`generate-excel-chart-families.ps1`):
// read from the package, resolved against live cells and painted by the shared chart painter.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { chartView, createRefEvaluator } from '../layout/index';
import { renderChartSvg } from '../../chart/render/chart-svg';
import { putCell } from '../cells';
import type { ChartObject, Workbook } from '../model';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-chart-families.xlsx')),
	);

const chartOf = (wb: Workbook, type: string): ChartObject => {
	for (const d of wb.sheets[0]!.drawings) {
		const chart = d.kind === 'chart' ? d : d.kind === 'unsupported' ? d.chartEx : undefined;
		if (chart?.chartType === type) return chart;
	}
	throw new Error(`no ${type} chart`);
};
const draw = (wb: Workbook, chart: ChartObject) =>
	renderChartSvg(chartView(wb, 0, chart, createRefEvaluator(wb)), 480, 300);

describe('Excel bubble chart', () => {
	it('reads bubble sizes and draws them from live cells', async () => {
		const wb = await loadXlsx(fixture());
		const bubble = chartOf(wb, 'bubble');
		expect(bubble.series[0]!.bubbleSizes).toEqual([5, 10, 3, 8]);
		expect(bubble.bubble).toMatchObject({ showNegativeBubbles: false });
		const model = chartView(wb, 0, bubble, createRefEvaluator(wb));
		expect(model.supported).toBe(true);
		expect(model.series[0]!.sizes).toEqual([5, 10, 3, 8]);
		const svg = draw(wb, bubble);
		expect(svg.match(/fill-opacity="0.6"/g)).toHaveLength(4);
		expect(svg).toContain('>Bubbles<');
		expect(svg).not.toContain('not drawn');
	});

	it('follows an edited bubble size', async () => {
		const wb = await loadXlsx(fixture());
		const bubble = chartOf(wb, 'bubble');
		// Sizes sit in column L (index 11), sheet rows 2 to 5 (0-based row 1).
		putCell(wb.sheets[0]!, 1, 11, { value: 40 });
		const model = chartView(wb, 0, bubble, createRefEvaluator(wb));
		expect(model.series[0]!.sizes![0]).toBe(40);
	});
});

describe('Excel stock chart', () => {
	it('draws open-high-low-close candles from live cells', async () => {
		const wb = await loadXlsx(fixture());
		const stock = chartOf(wb, 'stock');
		expect(stock.series.map((s) => s.name)).toEqual(['Open', 'High', 'Low', 'Close']);
		const svg = draw(wb, stock);
		const candles = svg.match(/<g data-chart-series="3"[\s\S]*?<\/g>/g)!;
		expect(candles).toHaveLength(4);
		// Closes 12, 13, 11, 13 against opens 10, 11, 12, 11: the third candle is a down bar.
		expect(candles.map((g) => /<rect [^>]*fill="(#[0-9A-F]+)"/.exec(g)![1])).toEqual([
			'#FFFFFF',
			'#FFFFFF',
			'#404040',
			'#FFFFFF',
		]);
		expect(svg).toContain('>Prices<');
	});
});

describe('Excel treemap and sunburst charts', () => {
	it('splits the two-column category range into a hierarchy and draws a cell per leaf', async () => {
		const wb = await loadXlsx(fixture());
		const treemap = chartOf(wb, 'treemap');
		const model = chartView(wb, 0, treemap, createRefEvaluator(wb));
		expect(model.categories).toEqual(['Apple', 'Pear', 'Leek', 'Kale', 'Almond']);
		expect(model.hierarchy!.branches).toEqual(['Fruit', 'Veg', 'Nuts']);
		const svg = draw(wb, treemap);
		expect(svg.match(/<g data-chart-series="0" data-chart-point/g)).toHaveLength(5);
		for (const label of ['Fruit', 'Veg', 'Nuts', 'Produce treemap'])
			expect(svg).toContain(`>${label}<`);
		expect(svg).not.toContain('not drawn');
	});

	it('draws the sunburst rings from the same cells', async () => {
		const wb = await loadXlsx(fixture());
		const svg = draw(wb, chartOf(wb, 'sunburst'));
		expect(svg.match(/<path d="M [^"]*" fill="/g)).toHaveLength(8);
		expect(svg).toContain('>Produce sunburst<');
	});

	it('follows an edited weight', async () => {
		const wb = await loadXlsx(fixture());
		// Weights sit in column V (index 21), rows 2 to 6.
		putCell(wb.sheets[0]!, 1, 21, { value: 99 });
		const model = chartView(wb, 0, chartOf(wb, 'treemap'), createRefEvaluator(wb));
		expect(model.series[0]!.values[0]).toBe(99);
	});

	it('saves every extended frame unchanged', async () => {
		const wb = await loadXlsx(fixture());
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of ['chartEx1', 'chartEx2', 'chartEx3', 'chartEx4'])
			expect(saved.file(`xl/charts/${part}.xml`)).not.toBeNull();
	});
});
