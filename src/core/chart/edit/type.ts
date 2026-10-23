// Changing the chart type of a whole chart (Change Chart Type). Semantics follow PowerPoint's
// save path for a type change (`PptxHandlerRuntimeSaveDataSerialization`, `chart-container-*`):
// the group keeps what its new element allows and drops the rest, gains the children the schema
// requires, demotes `clustered` grouping outside bar charts, and axes no group references any more
// are removed. Where PowerPoint has nothing to say Excel's choices apply (`xlsx/write/chart-space`):
// pie charts vary colours by point, a doughnut gets a 50% hole, a chart gaining axes gets Excel's
// two default axes, and stacked bars overlap fully. A combination chart becomes one group.
import { CHART_GROUP_ELEMENTS } from '../parse-plot';
import type { ChartAxis, ChartGroupKind, ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { assignOrDelete, chartSpaceSeries, editChartSpace, firstChartGroup } from './draft';
import { GROUP_FIELDS, groupHasAxes, stripSeriesForKind } from './type-schema';
import {
	DEFAULT_CHART_AXIS_IDS,
	convertChartAxis,
	defaultChartAxes,
	rotatedAxisPosition,
} from './type-axes';
import { isXYChartKind } from './series';
import { chartGroupPaintsLine, repaintSeriesForFamily } from './fill';

/** Grouping of bar, line and area charts. */
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

export interface ChartTypeTarget {
	kind: ChartGroupKind;
	/** Defaults to the current chart's 3-D state where the new family has a 3-D variant. */
	is3D?: boolean;
	/** Bar charts: `col` (vertical, the default) or `bar` (horizontal). */
	barDirection?: 'bar' | 'col';
	grouping?: ChartGrouping;
}

/** The families a whole chart can switch to here, and which have a 3-D variant. */
const SWITCHABLE: Readonly<Partial<Record<ChartGroupKind, boolean>>> = {
	bar: true,
	line: false,
	area: true,
	pie: true,
	doughnut: false,
	ofPie: false,
	scatter: false,
	radar: false,
	bubble: false,
};

function elementFor(kind: ChartGroupKind, is3D: boolean): string {
	const match = Object.entries(CHART_GROUP_ELEMENTS).find(
		([, value]) => value.kind === kind && value.is3D === is3D,
	);
	if (!match) throw new RangeError(`No ${is3D ? '3-D ' : ''}${kind} chart`);
	return match[0];
}

/** Legal grouping of `element`: `clustered` only in bar charts, `standard` in 3-D bars. */
function groupingFor(element: string, grouping: string | undefined): string {
	const bar = element === 'barChart' || element === 'bar3DChart';
	if (!bar) return !grouping || grouping === 'clustered' ? 'standard' : grouping;
	if (!grouping) return 'clustered';
	return grouping === 'standard' && element === 'barChart' ? 'clustered' : grouping;
}

const horizontalBars = (group: Pick<ChartPlotGroup, 'kind' | 'barDirection'>) =>
	group.kind === 'bar' && group.barDirection === 'bar';

function newGroup(base: ChartPlotGroup, target: ChartTypeTarget, series: ChartSpaceSeries[]) {
	const is3D = target.is3D ?? (base.is3D && SWITCHABLE[target.kind] === true);
	const element = elementFor(target.kind, is3D);
	const group: ChartPlotGroup = { kind: target.kind, element, is3D, series, axisIds: [] };
	const allowed = GROUP_FIELDS[element] ?? [];
	for (const key of allowed)
		if (base[key] !== undefined) Object.assign(group, { [key]: structuredClone(base[key]) });
	if (target.kind === 'bar') {
		group.barDirection = target.barDirection ?? base.barDirection ?? 'col';
		group.gapWidth ??= 150;
	}
	if (allowed.includes('grouping')) {
		group.grouping = groupingFor(element, target.grouping ?? base.grouping);
		if (target.kind === 'bar' && group.grouping !== 'clustered' && group.grouping !== 'standard')
			group.overlap ??= 100;
	}
	if (element !== 'bar3DChart') for (const item of series) delete item.shape;
	if (target.kind === 'line') group.marker ??= true;
	if (target.kind === 'pie' || target.kind === 'doughnut' || target.kind === 'ofPie')
		group.varyColors = true;
	if (target.kind === 'pie' || target.kind === 'doughnut') group.firstSliceAngle ??= 0;
	if (target.kind === 'doughnut') group.holeSize ??= 50;
	if (target.kind === 'ofPie') group.ofPieType ??= 'pie';
	if (target.kind === 'scatter') group.scatterStyle ??= 'lineMarker';
	if (target.kind === 'radar') group.radarStyle ??= 'marker';
	return group;
}

/** The axes of the new group: the old group's primary pair, converted, or Excel's defaults. */
function groupAxes(space: ChartSpace, base: ChartPlotGroup, group: ChartPlotGroup): ChartAxis[] {
	if (!groupHasAxes(group.element)) return [];
	const xy = isXYChartKind(group.kind);
	const byId = (id: number) => space.plotArea.axes.find((axis) => axis.id === id);
	const old = base.axisIds.map(byId).filter((axis): axis is ChartAxis => axis !== undefined);
	if (old.length < 2) {
		const taken = new Set(
			space.plotArea.axes.flatMap((axis) => (axis.id === undefined ? [] : [axis.id])),
		);
		const free = DEFAULT_CHART_AXIS_IDS.every((id) => !taken.has(id));
		const top = Math.max(0, ...taken);
		const ids: [number, number] = free ? [...DEFAULT_CHART_AXIS_IDS] : [top + 1, top + 2];
		return defaultChartAxes(horizontalBars(group), xy, ids);
	}
	const rotate = horizontalBars(base) !== horizontalBars(group);
	const axes = old
		.slice(0, group.element === 'line3DChart' || group.is3D ? 3 : 2)
		.map((axis, i) => {
			const kind = i === 0 ? (xy ? 'val' : axis.kind === 'val' ? 'cat' : axis.kind) : axis.kind;
			const next = convertChartAxis(axis, kind);
			if (rotate) assignOrDelete(next, 'position', rotatedAxisPosition(next.position));
			if (i === 1 && next.kind === 'val' && xy !== isXYChartKind(base.kind))
				next.crossBetween = xy ? 'midCat' : 'between';
			return next;
		});
	if (group.element === 'line3DChart' && axes.length < 3)
		throw new RangeError('A 3-D line chart needs a series axis the chart does not have');
	return axes;
}

/** Changes the type of the whole chart; series keep their data, names and formatting. */
export function setChartType(space: ChartSpace, target: ChartTypeTarget): ChartSpace {
	if (!Object.hasOwn(SWITCHABLE, target.kind))
		throw new RangeError(`Switching a chart to ${target.kind} is not supported`);
	return editChartSpace(space, (draft) => {
		const base = firstChartGroup(draft);
		const series = chartSpaceSeries(draft).map((item) =>
			stripSeriesForKind(item, base.kind, target.kind),
		);
		const group = newGroup(base, target, series);
		const toLine = chartGroupPaintsLine(group);
		if (toLine !== chartGroupPaintsLine(base))
			for (const item of series) repaintSeriesForFamily(item, toLine);
		if (
			group.is3D &&
			target.kind === 'bar' &&
			group.grouping === 'standard' &&
			base.axisIds.length < 3
		)
			group.grouping = 'clustered';
		const axes = groupAxes(draft, base, group);
		group.axisIds = axes.flatMap((axis) => (axis.id === undefined ? [] : [axis.id]));
		draft.plotArea.groups = [group];
		draft.plotArea.axes = axes;
		if (!group.is3D) {
			delete draft.view3D;
			delete draft.floor;
			delete draft.sideWall;
			delete draft.backWall;
		} else
			draft.view3D ??=
				target.kind === 'pie'
					? { rotX: 30, rotY: 0, rightAngleAxes: false }
					: { rotX: 15, rotY: 20, rightAngleAxes: true };
	});
}

/** Sets the grouping of every bar, line and area group (PowerPoint's `setChartGrouping`). */
export function setChartGrouping(space: ChartSpace, grouping: ChartGrouping): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const group of draft.plotArea.groups) {
			if (!(GROUP_FIELDS[group.element] ?? []).includes('grouping')) continue;
			group.grouping = groupingFor(group.element, grouping);
			if (group.kind !== 'bar') continue;
			if (group.grouping === 'clustered' || group.grouping === 'standard') {
				if (group.overlap === 100) delete group.overlap;
			} else group.overlap = 100;
		}
	});
}
