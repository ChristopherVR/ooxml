/**
 * Writes the part of a NEW classic chart (SDK `addChart`, `createChartElement`,
 * `ChartBuilder`, Insert Chart) through the shared chart core, as extended
 * charts already are (`chart-ex-neutral-create.ts`): `newChartSpace` builds the
 * part the legacy generator (`chart-xml-generator.ts`) wrote for the model's
 * type, with its data written as literals (no workbook) and its visible
 * defaults (legend at the right, gap width 150 with no overlap, doughnut hole
 * 50, no rounded corners, the generator's axis ids); a baseline describes what
 * that part holds, and the loaded-chart save planner (`planNeutralChartSave`)
 * turns every other field of the model into `chart/edit` operations, which
 * `writeChartSpace` writes.
 *
 * Rules for inputs the model can state twice:
 * - `style.hasTitle: false` removes the title, whatever `title` and `titleRuns`
 *   hold;
 * - non-empty `titleRuns` are the title; `title` is then ignored;
 * - `clustered` grouping of a line or area chart is written as `standard`, the
 *   only lower grouping their schema type has;
 * - a data label position the chart type does not offer is left out and
 *   reported (`chart-label-positions.ts`).
 *
 * Anything the shared core cannot write leaves `xml` undefined with the fields
 * that need the legacy generator, which then writes the part as before.
 *
 * @module utils/chart-new-part
 */
import {
	chartIssuesLostOnWrite,
	newChartSpace,
	parseChartSpace,
	setChartSeriesSecondaryAxis,
	writeChartSpace,
	type ChartGrouping,
	type ChartSpace,
	type ChartTypeTarget,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import { newChartAxes } from './chart-new-part-axes';
import { newChartBaseline, normalizedNewChart } from './chart-new-part-model';
import { newChartGroupSettings } from './chart-new-part-group';
import { planWithLeftovers } from './chart-neutral-save-plan';
import { NEUTRAL_CHART_TYPE_TARGETS } from './chart-neutral-save-type';
import { GENERATED_CHART_AXIS_IDS } from './chart-xml-container-map';

/** The part of a new chart, or why the legacy generator must write it. */
export interface NewChartPart {
	/** The `c:chartSpace` part written through the shared core; absent for the legacy generator. */
	xml?: string;
	/**
	 * Model fields, as paths, the written part leaves out (with `xml`), or that need the legacy
	 * generator, which may drop them (without `xml`).
	 */
	unwritten: string[];
	/**
	 * Without `xml`: the model the legacy generator writes, the chart with the label positions its
	 * type does not offer left out.
	 */
	legacyModel?: PptxChartData;
}

/** Model fields a new part leaves out: drawings and workbooks it would need parts for. */
const OUTSIDE_PART = [
	'userShapes',
	'userShapesXml',
	'externalData',
	'embeddedWorkbookData',
] as const satisfies readonly (keyof PptxChartData)[];

function targetOf(chart: PptxChartData): ChartTypeTarget | undefined {
	if (chart.chartType === 'combo') return NEUTRAL_CHART_TYPE_TARGETS['bar'];
	if (!Object.hasOwn(NEUTRAL_CHART_TYPE_TARGETS, chart.chartType)) return undefined;
	const target = NEUTRAL_CHART_TYPE_TARGETS[chart.chartType]!;
	return {
		...target,
		...(chart.chartType === 'surface' ? { is3D: chart.surfaceTopView === false } : {}),
		...(target.kind === 'bar' && chart.barDirection ? { barDirection: chart.barDirection } : {}),
		...(chart.grouping === undefined
			? {}
			: {
					grouping: (chart.groupingStandard && chart.grouping === 'clustered'
						? 'standard'
						: chart.grouping) as ChartGrouping,
				}),
	};
}

/** The part before the planner's edits: type, data, title, legend, group settings and axes. */
function startingPart(chart: PptxChartData, axes: ReturnType<typeof newChartAxes>): ChartSpace {
	const target = targetOf(chart)!;
	let space = newChartSpace(
		{
			categories: chart.categories,
			series: chart.series.map((series) => ({
				name: series.name,
				values: series.values.map((value, index) => (series.blanks?.[index] ? null : value)),
			})),
		},
		target,
		{
			data: 'literal',
			profile: 'minimal',
			legendPosition: chart.style?.hasLegend ? ((chart.style.legendPosition ?? 'r') as 'r') : null,
			axisIds: [
				GENERATED_CHART_AXIS_IDS.cat,
				GENERATED_CHART_AXIS_IDS.val,
				GENERATED_CHART_AXIS_IDS.ser,
			],
			dateAxis: axes.dateAxis,
			group: newChartGroupSettings(chart, target),
			...(chart.title === undefined ? {} : { title: chart.title }),
		},
	);
	if (chart.date1904 !== undefined) space = { ...space, date1904: chart.date1904 };
	for (const { seriesIndex, axisIds } of axes.secondary)
		space = setChartSeriesSecondaryAxis(space, seriesIndex, true, { axisIds });
	return space;
}

function writeThroughCore(chart: PptxChartData): NewChartPart {
	const legacy = (reason: string[]): NewChartPart => {
		const { after, unwritten } = normalizedNewChart(chart, { labelPositionsOnly: true });
		return { unwritten: [...unwritten, ...reason], legacyModel: after };
	};
	if (!targetOf(chart)) return legacy([`chartType.${chart.chartType}`]);
	if (chart.pivotSource) return legacy(['pivotSource']);
	const { after, unwritten } = normalizedNewChart(chart);
	const axes = newChartAxes(after);
	if (after.axes !== undefined || axes.after.length > 0) after.axes = axes.after;
	if (axes.elementFormats) after.elementFormats = axes.elementFormats;
	const { plan, unclaimed } = planWithLeftovers(newChartBaseline(after, axes), after);
	if (!plan) return legacy(unclaimed);
	let space = startingPart(after, axes);
	for (const edit of plan.edits) space = edit(space);
	const xml = writeChartSpace(space);
	const lost = chartIssuesLostOnWrite(parseChartSpace(xml).issues);
	if (lost.length > 0)
		return { unwritten: lost.map((issue) => `part.${issue.element ?? issue.code}`) };
	return {
		xml,
		unwritten: [
			...unwritten,
			...axes.unwritten,
			...OUTSIDE_PART.filter((key) => chart[key] !== undefined),
		],
	};
}

/** The part of the new chart `chart`, written through the shared chart core when it can be. */
export function writeNewChartPart(chart: PptxChartData): NewChartPart {
	try {
		return writeThroughCore(chart);
	} catch (error) {
		// A value a shared operation refuses (a combination its families cannot hold, a setting
		// outside its schema range) keeps the legacy generator, which writes the part as before.
		const { after, unwritten } = normalizedNewChart(chart, { labelPositionsOnly: true });
		return {
			unwritten: [
				...unwritten,
				`refused: ${error instanceof Error ? error.message : String(error)}`,
			],
			legacyModel: after,
		};
	}
}
