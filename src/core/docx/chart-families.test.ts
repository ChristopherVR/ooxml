// Combination, waterfall and histogram charts authored by Word (`generate-word-chart-families.ps1`):
// the waterfall and histogram are extended (`cx:`) parts with cached points, painted by the shared
// chart painter like the classic charts.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { chartsIn, loadDocx, type DocxChart } from './index';
import { docxChartPaint } from './chart-paint';
import { adaptDocumentModel } from './layout/adapter';

const bytes = () =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '__fixtures__/chart-families.docx')));

async function load() {
	const loaded = await loadDocx(bytes());
	const charts = chartsIn(loaded.model.blocks);
	const by = (title: string): DocxChart => charts.find((c) => c.title === title)!;
	return { loaded, charts, by };
}

describe('Word combination, waterfall and histogram charts', () => {
	it('finds every chart, reading the extended parts through their relationship', async () => {
		const { charts, by } = await load();
		expect(charts).toHaveLength(7);
		expect(by('Revenue and margin').chartSpace).toBeDefined();
		expect(by('Cash flow')).toMatchObject({ layoutId: 'waterfall' });
		expect(by('Cash flow').chartSpace).toBeUndefined();
		expect(by('Scores')).toMatchObject({ layoutId: 'clusteredColumn' });
		for (const chart of charts) expect(chart.issues).toEqual([]);
	});

	it('draws the combination with its line on a secondary axis', async () => {
		const { loaded, by } = await load();
		const paint = docxChartPaint(by('Revenue and margin'), 576, 336, loaded.model.theme);
		expect(paint.label).toBe('Chart: Revenue and margin');
		const svg = paint.svg!;
		expect(svg.match(/<g data-chart-series="0" data-chart-point/g)).toHaveLength(4);
		expect(svg).toContain('<polyline');
		expect(svg).toContain('>Margin<');
		expect(svg).toContain('>Revenue<');
		// The secondary axis scale is the margin fraction, not the revenue.
		expect(svg).toContain('>0.3<');
	});

	it('draws the waterfall from its cached points', async () => {
		const { loaded, by } = await load();
		const chart = by('Cash flow');
		expect(chart.chartEx?.waterfall?.subtotalIndices).toEqual([0, 5]);
		const paint = docxChartPaint(chart, 576, 336, loaded.model.theme);
		expect(paint.label).toBe('Chart: Cash flow');
		const svg = paint.svg!;
		expect(svg.match(/<g data-chart-series="0" data-chart-point/g)).toHaveLength(6);
		for (const label of ['Start', 'Sales', 'Returns', 'Costs', 'Other', 'Total'])
			expect(svg).toContain(`>${label}<`);
		for (const label of ['Increase', 'Decrease', 'Total']) expect(svg).toContain(`>${label}<`);
		expect(svg).not.toContain('not drawn');
	});

	it('draws the histogram from its cached observations', async () => {
		const { loaded, by } = await load();
		const paint = docxChartPaint(by('Scores'), 576, 336, loaded.model.theme);
		expect(paint.label).toBe('Chart: Scores');
		expect(paint.svg!.match(/<g data-chart-series="0" data-chart-point/g)!.length).toBeGreaterThan(
			1,
		);
	});

	it('shows a labelled frame for an extended family it does not draw', async () => {
		const { by } = await load();
		const funnel: DocxChart = { ...by('Cash flow'), layoutId: 'funnel' };
		delete funnel.chartEx;
		expect(docxChartPaint(funnel, 100, 100)).toEqual({ label: 'Chart (funnel)' });
	});

	it('carries the seven painted charts to the print layout input', async () => {
		const { loaded } = await load();
		const input = adaptDocumentModel(loaded.model);
		expect(JSON.stringify(input).match(/"chart":\{"svg":"<svg /g)).toHaveLength(7);
	});

	it('keeps the extended chart parts and drawings on save', async () => {
		const { loaded } = await load();
		const { saveDocx } = await import('./index');
		const saved = await JSZip.loadAsync(await saveDocx(loaded.model));
		expect(saved.file('word/charts/chartEx1.xml')).not.toBeNull();
		expect(saved.file('word/charts/chartEx2.xml')).not.toBeNull();
		const document = await saved.file('word/document.xml')!.async('string');
		expect(document.match(/drawing\/2014\/chartex/g)!.length).toBeGreaterThanOrEqual(2);
	});
});
