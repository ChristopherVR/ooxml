import JSZip from 'jszip';

import { ChartBuilder } from '../../core/builders/sdk/ChartBuilder';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import type { ChartInput } from '../../core/builders/sdk/types';
import { PptxHandler } from '../../core/PptxHandler';
import type {
	ChartPptxElement,
	PptxChartData,
	PptxChartType,
	PptxCompatibilityWarning,
} from '../../core/types';

/** How the chart is created: `SlideBuilder.addChart` or the fluent `ChartBuilder`. */
export type NewChartPath = 'addChart' | 'builder';

/** The sample data of a new chart: three series over three categories. */
export const NEW_CHART_INPUT: ChartInput = {
	series: [
		{ name: 'Revenue', values: [10, 20, 30] },
		{ name: 'Margin', values: [3, 4, 5] },
		{ name: 'Cost', values: [7, 16, 25] },
	],
	categories: ['Q1', 'Q2', 'Q3'],
};

export interface SavedNewChart {
	/** The saved chart part. */
	xml: string;
	/** The chart model after `new PptxHandler().load` of the saved bytes. */
	chart: PptxChartData;
	/** The save warnings about the chart. */
	warnings: PptxCompatibilityWarning[];
	bytes: Uint8Array;
}

/** A deck with one new chart, built through the public SDK, and a save that reloads it. */
export async function newChartDeck(
	type: PptxChartType = 'bar',
	input: Partial<ChartInput> = {},
	path: NewChartPath = 'addChart',
): Promise<{ element: ChartPptxElement; save: () => Promise<SavedNewChart> }> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const full: ChartInput = { ...NEW_CHART_INPUT, ...input };
	const position = { x: 50, y: 50, width: 500, height: 300 };
	const slide =
		path === 'addChart'
			? createSlide('Blank').addChart(type, full, position).build()
			: createSlide('Blank').build();
	if (path === 'builder') {
		let builder = ChartBuilder.create(type).categories(full.categories);
		for (const series of full.series) builder = builder.addSeries(series.name, series.values);
		if (full.title !== undefined) builder = builder.title(full.title);
		if (full.grouping !== undefined) builder = builder.grouping(full.grouping);
		slide.elements.push(
			builder.bounds(position.x, position.y, position.width, position.height).build(),
		);
	}
	data.slides.push(slide);
	const element = slide.elements.find((item) => item.type === 'chart') as ChartPptxElement;
	return {
		element,
		save: async () => {
			const bytes = await handler.save(data.slides);
			const xml = await (
				await JSZip.loadAsync(bytes)
			)
				.file('ppt/charts/chart1.xml')!
				.async('string');
			const warnings = handler
				.getCompatibilityWarnings()
				.filter((warning) => warning.code.startsWith('CHART_'));
			const loaded = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
			const chart = loaded.slides[0]!.elements.find((item) => item.type === 'chart') as
				| ChartPptxElement
				| undefined;
			return { xml, chart: chart!.chartData!, warnings, bytes };
		},
	};
}

/** The text of the first element named `local` in `xml` (`c:ser`, `c:valAx`...), with its tags. */
export function firstElement(xml: string, local: string, from = 0): string {
	const start = xml.indexOf(`<${local}>`, from);
	if (start < 0) return '';
	return xml.slice(start, xml.indexOf(`</${local}>`, start) + local.length + 3);
}
