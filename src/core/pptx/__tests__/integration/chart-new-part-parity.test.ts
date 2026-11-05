/**
 * A new SDK chart written through the shared chart core (`writeNewChartPart`) reloads as the same
 * chart as the legacy generator's part (`buildChartSpaceXml`) for every classic type: the deck is
 * saved once, then the legacy part is swapped in for a second load, and the two models compared.
 *
 * Intentional differences, asserted where they show: the axes of horizontal bars sit at the left
 * (category) and bottom (value) as Office places them, where the generator wrote bottom and left
 * for every chart.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import type { ChartInput } from '../../core/builders/sdk/types';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxChartType } from '../../core/types';
import { writeNewChartPart } from '../../core/utils/chart-new-part';
import { buildChartSpaceXml } from '../../core/utils/chart-xml-generator';

const builder = new PptxRuntimeDependencyFactory().createBuilder();

interface Case {
	name: string;
	type: PptxChartType;
	input?: Partial<ChartInput>;
	edit?: (data: PptxChartData) => void;
	seriesCount?: number;
}

const CASES: Case[] = [
	{ name: 'clustered column', type: 'bar' },
	{ name: 'stacked column', type: 'bar', input: { grouping: 'stacked' } },
	{ name: '100% stacked column', type: 'bar', input: { grouping: 'percentStacked' } },
	{ name: 'clustered bar', type: 'bar', input: { barDirection: 'bar' } },
	{
		name: 'stacked bar with spacing',
		type: 'bar',
		input: { barDirection: 'bar', grouping: 'stacked', gapWidth: 80, overlap: 100 },
	},
	{ name: 'line', type: 'line' },
	{ name: 'stacked line', type: 'line', input: { grouping: 'stacked' } },
	{ name: 'area', type: 'area' },
	{ name: '100% stacked area', type: 'area', input: { grouping: 'percentStacked' } },
	{ name: 'pie', type: 'pie', input: { firstSliceAngle: 45 }, seriesCount: 1 },
	{ name: 'doughnut', type: 'doughnut', seriesCount: 1 },
	{ name: 'doughnut with a hole', type: 'doughnut', input: { holeSize: 60 }, seriesCount: 1 },
	{ name: 'pie of pie', type: 'ofPie', seriesCount: 1 },
	{ name: 'scatter', type: 'scatter' },
	{ name: 'bubble', type: 'bubble' },
	{ name: 'radar', type: 'radar' },
	{ name: '3-D column', type: 'bar3D' },
	{ name: '3-D line', type: 'line3D' },
	{ name: '3-D pie', type: 'pie3D', seriesCount: 1 },
	{ name: '3-D area', type: 'area3D' },
	{ name: 'surface', type: 'surface' },
	{
		name: '3-D surface',
		type: 'surface',
		edit: (data) => {
			data.surfaceTopView = false;
		},
	},
	{ name: 'stock (high, low, close)', type: 'stock', seriesCount: 3 },
	{ name: 'stock (open, high, low, close)', type: 'stock', seriesCount: 4 },
];

const SERIES = [
	{ name: 'North', values: [4, 5, 6, 7], color: '#1F77B4' },
	{ name: 'South', values: [2, 3, 1, 4] },
	{ name: 'East', values: [6, 1, 3, 2], color: '#2ca02c' },
	{ name: 'West', values: [3, 4, 5, 6] },
];

/** Saves `item` through the handler and returns the deck with the part both ways. */
async function decks(item: Case): Promise<{ core: Uint8Array; legacy: Uint8Array; xml: string }> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addChart(
			item.type,
			{
				series: SERIES.slice(0, item.seriesCount ?? 3),
				categories: ['Q1', 'Q2', 'Q3', 'Q4'],
				title: 'Sales',
				...item.input,
			},
			{ x: 40, y: 40, width: 600, height: 360 },
		)
		.build();
	const chart = slide.elements.find((element) => element.type === 'chart') as ChartPptxElement;
	item.edit?.(chart.chartData!);
	data.slides.push(slide);
	const part = writeNewChartPart(structuredClone(chart.chartData!));
	expect(part.unwritten).toStrictEqual([]);
	expect(part.xml).toBeDefined();
	const core = await handler.save(data.slides);
	const zip = await JSZip.loadAsync(core);
	const xml = await zip.file('ppt/charts/chart1.xml')!.async('string');
	zip.file('ppt/charts/chart1.xml', builder.build(buildChartSpaceXml(chart.chartData!)));
	return { core, legacy: await zip.generateAsync({ type: 'uint8array' }), xml };
}

async function reload(bytes: Uint8Array): Promise<PptxChartData> {
	const loaded = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
	const chart = loaded.slides[0]!.elements.find((element) => element.type === 'chart');
	return (chart as ChartPptxElement).chartData!;
}

/** The fields both writers set, without the part path and parse-time bookkeeping. */
function comparable(chart: PptxChartData, horizontal: boolean) {
	return {
		chartType: chart.chartType,
		grouping: chart.grouping,
		barDirection: chart.barDirection,
		varyColors: chart.varyColors,
		categories: chart.categories,
		title: chart.title,
		titleRuns: chart.titleRuns,
		style: chart.style,
		series: chart.series.map((series) => ({
			name: series.name,
			values: series.values,
			color: series.color,
			seriesChartType: series.seriesChartType,
			xValues: series.xValues,
			bubbleSizes: series.bubbleSizes,
			blanks: series.blanks,
		})),
		axes: chart.axes?.map(({ axPos, ...axis }) => ({
			...axis,
			// Horizontal bars: the shared core places the axes where Office does.
			...(horizontal ? {} : { axPos }),
		})),
		options: [
			chart.barGapWidth,
			chart.barOverlap,
			chart.firstSliceAngle,
			chart.doughnutHoleSize,
			chart.ofPieOptions,
			chart.radarStyle,
			chart.scatterStyle,
			chart.wireframe,
			chart.surfaceTopView,
			chart.hiLowLines,
			chart.roundedCorners,
			chart.chartChrome,
		],
	};
}

describe('new classic charts: shared core against the legacy generator', () => {
	it.each(CASES)('$name reloads as the legacy part did', async (item) => {
		const { core, legacy, xml } = await decks(item);
		expect(xml).not.toContain('<c:f>');
		expect(xml).not.toContain('roundedCorners');
		expect(xml).toContain('<c:legendPos val="r"></c:legendPos>');
		const horizontal = item.input?.barDirection === 'bar';
		const [fromCore, fromLegacy] = [await reload(core), await reload(legacy)];
		expect(comparable(fromCore, horizontal)).toStrictEqual(comparable(fromLegacy, horizontal));
		if (horizontal) expect(fromCore.axes?.map((axis) => axis.axPos)).toStrictEqual(['l', 'b']);
	});
});
