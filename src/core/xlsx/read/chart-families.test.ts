// Combination, waterfall and histogram charts authored by Excel (`generate-excel-chart-families.ps1`):
// read from the package, resolved against live cells and painted by the shared chart painter.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { chartView, createRefEvaluator } from '../layout/index';
import { renderChartSvg } from '../../chart/render/chart-svg';
import { putCell } from '../cells';
import type { ChartObject, DrawingObject, Workbook } from '../model';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-chart-families.xlsx')),
	);

const drawings = (wb: Workbook): DrawingObject[] => wb.sheets[0]!.drawings;
const chartOf = (wb: Workbook, type: string): ChartObject => {
	for (const d of drawings(wb)) {
		const chart = d.kind === 'chart' ? d : d.kind === 'unsupported' ? d.chartEx : undefined;
		if (chart?.chartType === type) return chart;
	}
	throw new Error(`no ${type} chart`);
};
const draw = (wb: Workbook, chart: ChartObject) =>
	renderChartSvg(chartView(wb, 0, chart, createRefEvaluator(wb)), 480, 300);

describe('Excel combination chart', () => {
	it('keeps the line group on its secondary axis and draws both groups from live cells', async () => {
		const wb = await loadXlsx(fixture());
		const combo = chartOf(wb, 'column');
		expect(combo.series.map((s) => s.name)).toEqual(['Revenue']);
		expect(combo.extraGroups).toHaveLength(1);
		expect(combo.extraGroups![0]).toMatchObject({ chartType: 'line', secondary: true });
		const model = chartView(wb, 0, combo, createRefEvaluator(wb));
		expect(model.extraGroups![0]!.series[0]!.values).toEqual([0.12, 0.2, 0.15, 0.3]);
		expect(combo.axes?.length).toBeGreaterThanOrEqual(4);
		expect(model.axes?.value?.source.kind).toBe('val');
		expect(model.axes?.secondary?.source.position).toBe('r');
		expect(model.secondaryAxis!.max).toBeLessThanOrEqual(0.4);
		expect(model.valueAxis!.max).toBeGreaterThanOrEqual(190);
		const svg = draw(wb, combo);
		expect(svg.match(/<g data-chart-series="0" data-chart-point/g)).toHaveLength(4);
		expect(svg).toContain('<polyline');
		expect(svg).toContain('>Margin<');
	});

	it('follows an edited cell of the secondary series', async () => {
		const wb = await loadXlsx(fixture());
		putCell(wb.sheets[0]!, 4, 2, { value: 0.5 });
		const model = chartView(wb, 0, chartOf(wb, 'column'), createRefEvaluator(wb));
		expect(model.extraGroups![0]!.series[0]!.values[3]).toBe(0.5);
		expect(model.secondaryAxis!.max).toBeGreaterThanOrEqual(0.5);
	});
});

describe('Excel extended charts', () => {
	it('reads the waterfall through its hidden defined names and draws running totals', async () => {
		const wb = await loadXlsx(fixture());
		const frame = drawings(wb).find(
			(d) => d.kind === 'unsupported' && d.chartEx?.chartType === 'waterfall',
		);
		expect(frame).toMatchObject({ kind: 'unsupported', description: 'chart (waterfall)' });
		const waterfall = chartOf(wb, 'waterfall');
		expect(waterfall.waterfall?.subtotalIndices).toEqual([0, 5]);
		const model = chartView(wb, 0, waterfall, createRefEvaluator(wb));
		expect(model.categories).toEqual(['Start', 'Sales', 'Returns', 'Costs', 'Other', 'Total']);
		expect(model.waterfall!.steps.map((s) => s.endValue)).toEqual([100, 160, 140, 95, 110, 110]);
		const svg = draw(wb, waterfall);
		expect(svg).toContain('>Cash flow<');
		expect(svg.match(/<g data-chart-series="0" data-chart-point/g)).toHaveLength(6);
		expect(svg).not.toContain('not drawn');
	});

	it('reads the histogram and bins the live observations', async () => {
		const wb = await loadXlsx(fixture());
		const histogram = chartOf(wb, 'histogram');
		expect(histogram.histogram).toMatchObject({ intervalClosed: 'r' });
		const model = chartView(wb, 0, histogram, createRefEvaluator(wb));
		expect(model.series[0]!.values.reduce<number>((sum, v) => sum + (v ?? 0), 0)).toBe(20);
		expect(draw(wb, histogram)).toContain('>Scores<');
	});

	it('saves the extended chart frames unchanged', async () => {
		const wb = await loadXlsx(fixture());
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		const drawing = await saved.file('xl/drawings/drawing1.xml')!.async('string');
		expect(drawing.match(/chartex/g)!.length).toBeGreaterThanOrEqual(2);
		expect(saved.file('xl/charts/chartEx1.xml')).not.toBeNull();
		expect(saved.file('xl/charts/chartEx2.xml')).not.toBeNull();
	});
});
