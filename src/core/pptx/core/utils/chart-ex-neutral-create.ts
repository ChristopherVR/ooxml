// PowerPoint creation adapts view data to the neutral model, preserving the raw geography fallback.
import { isXmlNode } from './xml-access';
import { XMLParser } from 'fast-xml-parser';
import {
	createChartExSpace,
	setChartExDataLabelsVisible,
	writeChartExSpace,
	type ChartExFamily,
} from '../../../chart';
import type { PptxChartData, XmlObject } from '../types';
import { planChartExFamilyEdits } from './chart-ex-neutral-save-families';
import type { ChartExEdit } from './chart-ex-neutral-save';
/**
 * A pareto chart is written as Office writes it: the frequency series plus an owned `paretoLine`
 * series and a percentage axis, never the cumulative values the view model computes. A region map
 * without a provider cache or source point indexes is written through the neutral model too; one
 * that carries them keeps the raw generator, which preserves them.
 */
export function createChartExThroughNeutral(chart: PptxChartData): XmlObject | undefined {
	const layouts: Partial<Record<PptxChartData['chartType'], ChartExFamily>> = {
		funnel: 'funnel',
		waterfall: 'waterfall',
		histogram: 'clusteredColumn',
		boxWhisker: 'boxWhisker',
		treemap: 'treemap',
		sunburst: 'sunburst',
		regionMap: 'regionMap',
	};
	const pareto = chart.series.some((series) => series.histogramOptions?.layout === 'pareto');
	const layout = pareto && chart.chartType === 'histogram' ? 'pareto' : layouts[chart.chartType];
	const map = chart.series[0]?.regionMapOptions;
	if (
		!layout ||
		(layout === 'regionMap' &&
			(chart.series.length !== 1 ||
				map?.geographyCache !== undefined ||
				map?.geographyCacheXml !== undefined ||
				map?.entityIds?.length ||
				map?.categorySourceIndices ||
				map?.valueSourceIndices ||
				map?.valueColors?.length ||
				map?.fixedColor))
	)
		return undefined;
	const dataSeries = chart.series.filter((series) => series.histogramOptions?.layout !== 'pareto');
	let space = createChartExSpace(
		{
			categories: chart.categories,
			categoryLevels: chart.categoryLevels ?? [chart.categories],
			series: dataSeries.map((series) => ({
				name: series.name,
				values: series.values.map((value, index) => (series.blanks?.[index] ? null : value)),
			})),
		},
		layout,
		{
			...(layout === 'regionMap'
				? {
						geography: {
							...(map?.projectionType ? { projectionType: map.projectionType } : {}),
							...(map?.viewedRegionType ? { viewedRegionType: map.viewedRegionType } : {}),
							...(map?.cultureLanguage ? { cultureLanguage: map.cultureLanguage } : {}),
							...(map?.cultureRegion ? { cultureRegion: map.cultureRegion } : {}),
							...(map?.attribution !== undefined ? { attribution: map.attribution } : {}),
						},
					}
				: {}),
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
