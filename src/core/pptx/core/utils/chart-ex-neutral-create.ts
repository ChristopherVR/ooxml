// PowerPoint creation adapts view data to the neutral model, preserving the raw geography fallback.
import { isXmlNode } from './xml-access';
import { XMLParser } from 'fast-xml-parser';
import {
	createChartExSpace,
	setChartExDataLabelsVisible,
	writeChartExSpace,
	type ChartExLayout,
} from '../../../chart';
import type { PptxChartData, XmlObject } from '../types';
import { planChartExFamilyEdits } from './chart-ex-neutral-save-families';
import type { ChartExEdit } from './chart-ex-neutral-save';
export function createChartExThroughNeutral(chart: PptxChartData): XmlObject | undefined {
	const layouts: Partial<Record<PptxChartData['chartType'], Exclude<ChartExLayout, 'regionMap'>>> =
		{
			funnel: 'funnel',
			waterfall: 'waterfall',
			histogram: 'clusteredColumn',
			boxWhisker: 'boxWhisker',
			treemap: 'treemap',
			sunburst: 'sunburst',
		};
	const layout = layouts[chart.chartType];
	if (!layout || chart.series.some((series) => series.histogramOptions?.layout === 'pareto'))
		return undefined;
	let space = createChartExSpace(
		{
			categories: chart.categories,
			categoryLevels: chart.categoryLevels ?? [chart.categories],
			series: chart.series.map((series) => ({
				name: series.name,
				values: series.values.map((value, index) => (series.blanks?.[index] ? null : value)),
			})),
		},
		layout,
		{
			...(chart.title ? { title: chart.title } : {}),
			...(chart.style?.hasLegend
				? { legend: { visible: true, position: chart.style.legendPosition } }
				: {}),
		},
	);
	const baseline = structuredClone(chart);
	baseline.series.forEach((series) => {
		delete series.color;
		delete series.waterfallOptions;
		delete series.histogramOptions;
		delete series.boxWhiskerOptions;
		delete series.treemapOptions;
	});
	const edits: ChartExEdit[] = [];
	planChartExFamilyEdits(baseline, structuredClone(chart), edits);
	for (const edit of edits) space = edit(space);
	if (chart.style?.hasDataLabels) {
		// The shared setter creates the schema-shaped visibility container.
		space = setChartExDataLabelsVisible(space, true);
	}
	const tree = new XMLParser({
		ignoreAttributes: false,
		parseTagValue: false,
		trimValues: false,
		isArray: (name) => name === 'cx:series' || name === 'cx:data',
	}).parse(writeChartExSpace(space)) as XmlObject;
	const root = tree['cx:chartSpace'] as XmlObject;
	if (isXmlNode(chart.userShapesXml)) {
		root['@_xmlns:c'] = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
		root['c:userShapes'] = chart.userShapesXml;
	}
	return tree;
}
