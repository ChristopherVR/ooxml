// The `minimal` profile of a new chart part (`new-chart.ts`): the group and axes of the picked type
// with only what the schema requires (ECMA-376 Part 1, 21.2.2: CT_BarChart's `c:barDir`,
// CT_LineChart's `c:grouping`, CT_ScatterChart's `c:scatterStyle`, CT_RadarChart's `c:radarStyle`,
// CT_OfPieChart's `c:ofPieType`, every axis's id, scaling, position and crossing axis), plus
// `c:varyColors` and `c:delete`, whose absence Office reads differently from a written default.
// Nothing Office adds on its own (gap widths, overlaps, markers, gridlines, tick marks, a 3-D view,
// the title's paragraph and run properties) is written: a host states its own defaults through the
// group settings and the edit operations.
import type { ChartAxis, ChartAxisKind, ChartPlotGroup } from '../model';
import type { ChartSpaceSeries, ChartTitle } from '../model-series';
import { chartTextBodyFromXml } from '../write-shape';
import { escapeText } from '../write-util';
import { isXYChartKind } from './series';
import { chartGroupElement, type ChartTypeTarget } from './type';
import { GROUP_FIELDS, groupHasAxes, stripSeriesForKind } from './type-schema';

const VARIED = new Set(['pie', 'doughnut', 'ofPie']);

/** The grouping the schema allows `element`: `clustered` is a bar grouping only (ST_Grouping). */
function schemaGrouping(element: string, grouping: string | undefined): string {
	const bar = element === 'barChart' || element === 'bar3DChart';
	if (bar) return grouping ?? 'clustered';
	return grouping === undefined || grouping === 'clustered' ? 'standard' : grouping;
}

/** The chart group of a `minimal` new chart showing `series` as `target`. */
export function minimalChartGroup(
	target: ChartTypeTarget,
	series: ChartSpaceSeries[],
): ChartPlotGroup {
	const is3D = target.is3D ?? false;
	const element = chartGroupElement(target.kind, is3D);
	const allowed = GROUP_FIELDS[element] ?? [];
	// A scatter or bubble series plots its categories as x values (`c:xVal`, `c:yVal`).
	const own = isXYChartKind(target.kind)
		? series.map((item) => stripSeriesForKind(item, 'bar', target.kind))
		: series;
	const group: ChartPlotGroup = { kind: target.kind, element, is3D, series: own, axisIds: [] };
	if (target.kind === 'bar') group.barDirection = target.barDirection ?? 'col';
	if (allowed.includes('grouping')) group.grouping = schemaGrouping(element, target.grouping);
	if (allowed.includes('varyColors')) group.varyColors = VARIED.has(target.kind);
	if (target.kind === 'ofPie') group.ofPieType = 'pie';
	if (target.kind === 'scatter') group.scatterStyle = 'lineMarker';
	if (target.kind === 'radar') group.radarStyle = 'marker';
	return group;
}

function axis(kind: ChartAxisKind, id: number, crossAxisId: number, position: string): ChartAxis {
	return {
		kind,
		id,
		crossAxisId,
		position,
		scaling: { orientation: 'minMax' },
		deleted: false,
		majorGridlines: false,
		minorGridlines: false,
	};
}

/**
 * The axes of a `minimal` new chart `group`, with `ids` for the category (or X value), value (or
 * Y value) and series axes: none for the pie family, a series axis for a 3-D line or a surface.
 */
export function minimalChartAxes(
	group: ChartPlotGroup,
	ids: readonly [number, number, number],
): ChartAxis[] {
	if (!groupHasAxes(group.element)) return [];
	const [first, second, third] = ids;
	const horizontal = group.kind === 'bar' && group.barDirection === 'bar';
	const axes = [
		axis(isXYChartKind(group.kind) ? 'val' : 'cat', first, second, horizontal ? 'l' : 'b'),
		axis('val', second, first, horizontal ? 'b' : 'l'),
	];
	if (group.element === 'line3DChart' || group.kind === 'surface')
		axes.push(axis('ser', third, second, 'b'));
	return axes;
}

/** A one-run title with no paragraph or run properties (`a:bodyPr`, `a:lstStyle`, `a:p/a:r`). */
export function minimalChartTitle(text: string): ChartTitle {
	const rich = chartTextBodyFromXml(
		`<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeText(text)}</a:t></a:r></a:p>`,
	);
	return { tx: { rich, text: rich.text }, text: rich.text, overlay: false };
}
