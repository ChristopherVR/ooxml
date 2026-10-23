// An xlsx `ChartPatch` as `chart/edit` operations on the chart's `ChartSpace` (chart step C2):
// `updateChart` keeps its summary-shaped API, and this module turns the difference between the
// summary the chart showed and the patched summary into edits of the model the part is saved
// from. Order: chart type and grouping, Quick Layout, title, legend, bar spacing, series, axes,
// element fills. Semantics follow the DOM patch path this replaces (`write/chart-patch.ts`, removed):
// a title edit keeps the title's formatting, a new legend is written as Excel writes one, series
// colours land where the summary reads them; a type change now keeps the chart's formatting
// (`setChartType`) where the old path regenerated the part from the summary. Families the type
// switch does not build (stock, surface, the chartex families) still regenerate.
// Not mapped (the summary fields are display-only here, as before): `extraGroups`, `bubble`,
// `stock`, `waterfall`, `histogram`, `hierarchy`, `axisIds` and series data labels, trendlines and
// error bars.
import type { ChartSpace } from '../../chart/model';
import { editChartSpace } from '../../chart/edit/draft';
import { setChartElementFill } from '../../chart/edit/fill';
import { setChartLegend } from '../../chart/edit/legend-labels';
import { applyChartQuickLayout, chartSupportsQuickLayout } from '../../chart/edit/quick-layout';
import { setChartTitle } from '../../chart/edit/title';
import { setChartGrouping, setChartType, type ChartTypeTarget } from '../../chart/edit/type';
import type { ChartObject } from '../model';
import { newChartSpaceOf } from '../write/chart';
import type { ChartElementFills } from './chart-element-fill';
import { chartSpaceWithSeries } from './chart-space-series';

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** The `setChartType` target of a summary chart type, or `undefined` where it builds none. */
function typeTarget(chart: ChartObject): ChartTypeTarget | undefined {
	const grouping = chart.grouping ? { grouping: chart.grouping } : {};
	switch (chart.chartType) {
		case 'column':
		case 'bar':
			return { kind: 'bar', barDirection: chart.chartType === 'bar' ? 'bar' : 'col', ...grouping };
		case 'line':
		case 'area':
			return { kind: chart.chartType, ...grouping };
		case 'pie':
		case 'doughnut':
		case 'scatter':
		case 'radar':
		case 'bubble':
			return { kind: chart.chartType };
		default:
			return undefined;
	}
}

/** The chart type and grouping of `after`, or `undefined` when the switch must regenerate. */
function withType(space: ChartSpace, before: ChartObject, after: ChartObject) {
	const typeChanged = after.chartType !== before.chartType;
	const groupingChanged = after.grouping !== undefined && after.grouping !== before.grouping;
	if (!typeChanged && !groupingChanged) return space;
	const target = typeTarget(after);
	if (!target || !space.plotArea.groups[0]) return undefined;
	try {
		return typeChanged
			? setChartType(space, target)
			: setChartGrouping(space, after.grouping as NonNullable<ChartObject['grouping']>);
	} catch {
		return undefined;
	}
}

/** Bar gap width and overlap of the first bar group (`undefined` removes the setting). */
function withSpacing(space: ChartSpace, before: ChartObject, after: ChartObject): ChartSpace {
	if (before.barGapWidth === after.barGapWidth && before.barOverlap === after.barOverlap)
		return space;
	return editChartSpace(space, (draft) => {
		const group = draft.plotArea.groups.find((item) => item.kind === 'bar');
		if (!group) return;
		if (after.barGapWidth === undefined) delete group.gapWidth;
		else group.gapWidth = after.barGapWidth;
		if (after.barOverlap === undefined) delete group.overlap;
		else group.overlap = after.barOverlap;
	});
}

function withLegend(space: ChartSpace, before: ChartObject, after: ChartObject): ChartSpace {
	if (!after.showLegend)
		return before.showLegend ? setChartLegend(space, { visible: false }) : space;
	if (!space.legend) return setChartLegend(space, { position: after.legendPosition ?? 'r' });
	if (after.legendPosition && after.legendPosition !== before.legendPosition)
		return setChartLegend(space, { position: after.legendPosition });
	return space;
}

/**
 * `space` (the model `before` was derived from) with the edits of `after`, the patched summary.
 * `layout` is the Quick Layout the patch applies, `fills` its element fills. Returns `space` itself
 * when nothing changed.
 */
export function chartSpaceWithPatch(
	space: ChartSpace,
	before: ChartObject,
	after: ChartObject,
	options: { layout?: number; fills?: ChartElementFills } = {},
): ChartSpace {
	let next = withType(space, before, after);
	if (!next) return newChartSpaceOf(after);
	if (options.layout !== undefined && chartSupportsQuickLayout(next))
		next = applyChartQuickLayout(next, options.layout, ['dataLabels', 'axes', 'dataTable']);
	if (before.title !== after.title) next = setChartTitle(next, after.title);
	next = withLegend(next, before, after);
	next = withSpacing(next, before, after);
	next = chartSpaceWithSeries(next, before, after);
	if (options.layout === undefined && after.axes && !same(before.axes, after.axes)) {
		const axes = structuredClone(after.axes);
		next = editChartSpace(next, (draft) => {
			draft.plotArea.axes = axes;
		});
	}
	for (const [part, fill] of Object.entries(options.fills ?? {}))
		if ((part !== 'title' || next.title) && (part !== 'legend' || next.legend))
			next = setChartElementFill(next, part as keyof ChartElementFills, fill);
	return next;
}
