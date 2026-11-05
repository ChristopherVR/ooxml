import JSZip from 'jszip';
import { PptxHandler, PresentationBuilder } from 'ooxml-core/pptx';
import type { ChartPptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createDefaultChartElement, type InsertChartKind } from './insert-chart';

/** Saves a deck holding the Insert Chart default for `kind`, then reloads it. */
async function saveInserted(
	kind: InsertChartKind,
	options: Parameters<typeof createDefaultChartElement>[2] = {},
) {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	const element = createDefaultChartElement(kind, undefined, options);
	slide.elements.push(element);
	data.slides.push(slide);
	const bytes = await handler.save(data.slides);
	const xml = await (await JSZip.loadAsync(bytes)).file('ppt/charts/chart1.xml')!.async('string');
	const warnings = handler
		.getCompatibilityWarnings()
		.filter((warning) => warning.code.startsWith('CHART_'));
	const loaded = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
	const chart = loaded.slides[0]!.elements.find(
		(item): item is ChartPptxElement => item.type === 'chart',
	)!.chartData!;
	return { element, xml, chart, warnings };
}

describe('Insert Chart: saved through the shared chart core', () => {
	it.each(['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter'] as const)(
		'saves the %s default with the visible defaults and no warning',
		async (kind) => {
			const { xml, chart, warnings } = await saveInserted(kind);
			expect(warnings).toStrictEqual([]);
			expect(xml).not.toContain('<c:f>');
			expect(xml).not.toContain('roundedCorners');
			expect(xml).toContain('<c:legendPos val="r"></c:legendPos>');
			expect(chart.title).toBe('Chart Title');
			expect(chart.series.map((series) => [series.name, series.values])).toStrictEqual([
				['Series 1', [4, 3, 5]],
			]);
			if (kind === 'column' || kind === 'bar') {
				expect(xml).toContain('<c:gapWidth val="150"></c:gapWidth>');
				expect(xml).not.toContain('<c:overlap');
				expect(chart.barDirection).toBe(kind === 'bar' ? 'bar' : 'col');
			}
			if (kind === 'doughnut') expect(chart.doughnutHoleSize).toBe(50);
		},
	);

	it('writes a line chart the dialog asked to cluster as standard', async () => {
		const { xml, chart } = await saveInserted('line', { grouping: 'clustered' });
		expect(xml).toContain('<c:grouping val="standard"></c:grouping>');
		expect(chart.groupingStandard).toBe(true);
	});

	it('writes the stacked groupings, an empty title as none, and no legend', async () => {
		const stacked = await saveInserted('column', { grouping: 'percentStacked' });
		expect(stacked.chart.grouping).toBe('percentStacked');
		const bare = await saveInserted('area', { title: '  ', hasLegend: false });
		expect(bare.xml).not.toContain('<c:title>');
		expect(bare.xml).not.toContain('<c:legend>');
		expect(bare.chart.title).toBeUndefined();
	});
});
