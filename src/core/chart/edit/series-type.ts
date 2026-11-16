import type { ChartAxis, ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { setChartType, type ChartTypeTarget } from './type';
import { convertChartAxis, defaultChartAxes } from './type-axes';

/** Whether two groups differ only in their series, so one group can hold both. */
function sameGroupSettings(a: ChartPlotGroup, b: ChartPlotGroup): boolean {
	const { series: _a, ...first } = a;
	const { series: _b, ...second } = b;
	return JSON.stringify(first) === JSON.stringify(second);
}

/**
 * The plot area with every group that differs from an earlier one only in its series joined into
 * it. Office writes one group per type, settings and axis pair, and PowerPoint draws two column
 * groups on the same axes over each other, so a combination built series by series
 * (`setChartSeriesType`, `setChartSeriesSecondaryAxis`, which split groups to keep flattened series
 * positions) is joined before it is written. The flattened series order changes when groups that
 * were not neighbours join, so this runs after the edits that address series by position.
 */
export function joinChartGroups(space: ChartSpace): ChartSpace {
	return editChartSpace(space, (draft) => {
		const groups: ChartPlotGroup[] = [];
		for (const group of draft.plotArea.groups) {
			const same = groups.find((kept) => sameGroupSettings(kept, group));
			if (same) same.series.push(...group.series);
			else groups.push(group);
		}
		draft.plotArea.groups = groups;
	});
}

/**
 * The families one series can take in a combination. A scatter series in a category combination
 * shares the category and value axes of the group it leaves, with its categories as x values: the
 * structure Excel 16 saves when a column or line series becomes XY (Scatter) through its object
 * model (`Series.ChartType`), observed in the files it writes.
 */
const COMBO_KINDS: readonly string[] = ['bar', 'line', 'area', 'scatter'];

/**
 * A column, line or area series that becomes XY (Scatter) as Excel 16 saves it: markers in the
 * series colour (a 5 pt circle with a 0.75 pt outline), no connecting line and `c:smooth` off, since
 * PowerPoint draws a scatter series without `c:smooth` as smoothed lines (it reports such a series
 * as `xlXYScatterSmooth`). A marker or smoothing the series already has is kept.
 */
function markerOnlyScatter(series: ChartSpaceSeries, original: ChartSpaceSeries): void {
	series.smooth ??= false;
	const paint = original.spPr?.fill;
	if (series.spPr?.line) {
		const { fill: _fill, ...line } = series.spPr.line;
		series.spPr = { ...series.spPr, line: { ...line, fill: { kind: 'none' } } };
	} else series.spPr = { ...series.spPr, line: { fill: { kind: 'none' } } };
	if (series.marker || !paint || paint.kind === 'none') return;
	series.marker = {
		symbol: 'circle',
		size: 5,
		spPr: {
			fill: structuredClone(paint),
			...(paint.kind === 'solid' ? { line: { widthEmu: 9525, fill: structuredClone(paint) } } : {}),
		},
	};
}

/** Whether a group draws horizontal bars (its category axis on the left). */
const horizontal = (group: Pick<ChartPlotGroup, 'kind' | 'barDirection'>) =>
	group.kind === 'bar' && group.barDirection === 'bar';

/** The coherent category (or scatter x value) and value axis pair of `group`, or `undefined`. */
function axisPair(space: ChartSpace, group: ChartPlotGroup): [ChartAxis, ChartAxis] | undefined {
	const pair = group.axisIds.map((id) => space.plotArea.axes.find((axis) => axis.id === id));
	const [x, y] = pair;
	if (
		pair.length !== 2 ||
		!x ||
		!y ||
		y.kind !== 'val' ||
		x.crossAxisId !== y.id ||
		y.crossAxisId !== x.id ||
		!['cat', 'date', ...(group.kind === 'scatter' ? ['val'] : [])].includes(x.kind)
	)
		return undefined;
	return [x, y];
}

/**
 * The secondary axis pair Excel 16 adds when one series of a combination turns between horizontal
 * bars and a vertical family (observed in the files its object model saves for
 * `Series.ChartType`): a hidden category axis along the side the new group's categories run and a
 * value axis opposite the primary one, crossing at its maximum, without gridlines.
 */
function secondaryPair(space: ChartSpace, category: ChartAxis, toHorizontal: boolean): ChartAxis[] {
	const taken = new Set(
		space.plotArea.axes.flatMap((axis) => (axis.id === undefined ? [] : [axis.id])),
	);
	const allocate = () => {
		let id = 500000003;
		while (taken.has(id)) id++;
		if (id > 4294967295) throw new RangeError('No chart axis IDs remain');
		taken.add(id);
		return id;
	};
	const ids: [number, number] = [allocate(), allocate()];
	const axes = defaultChartAxes(toHorizontal, false, ids);
	const hidden: ChartAxis = { ...structuredClone(category), id: ids[0], crossAxisId: ids[1] };
	delete hidden.crosses;
	delete hidden.majorGridlinesSpPr;
	delete hidden.minorGridlinesSpPr;
	hidden.deleted = true;
	hidden.position = toHorizontal ? 'l' : 'b';
	hidden.majorGridlines = false;
	hidden.minorGridlines = false;
	const value = axes[1];
	value.position = toHorizontal ? 't' : 'r';
	value.crosses = 'max';
	value.majorGridlines = false;
	return [hidden, value];
}

/** Turns the x value axis of an XY pair into the category axis a category family plots on. */
function categoryFromValueAxis(x: ChartAxis, y: ChartAxis): void {
	const category = convertChartAxis(x, 'cat');
	for (const key of Object.keys(x) as (keyof ChartAxis)[]) delete x[key];
	Object.assign(x, category);
	y.crossBetween = 'between';
}

/**
 * Changes one series of a 2-D combination without changing document-order selections or axis ids.
 *
 * - A column, line or area series may take another of those families or XY (Scatter); a scatter
 *   series shares the category and value axes of the group it leaves, with its categories as x
 *   values, as Excel 16 saves it.
 * - A series of an XY chart (two value axes) that takes a column, line or area family turns the x
 *   value axis into a category axis the remaining scatter series share, as Excel 16 does.
 * - A series that turns between horizontal bars and a vertical family moves to a secondary axis
 *   pair, as Excel 16 does; one that turns back joins the group of its orientation, and a pair
 *   nothing plots on any more goes.
 * - The series joins a group of the type it takes on the same axes, taking that group's settings,
 *   so a series converted back rejoins the group it left (Excel 16 writes one group again);
 *   neighbouring groups that then differ only in their series are joined.
 */
export function setChartSeriesType(
	space: ChartSpace,
	index: number,
	target: ChartTypeTarget,
): ChartSpace {
	const { group, position } = chartSeriesEntryAt(space, index);
	if (
		group.kind === target.kind &&
		group.is3D === (target.is3D ?? false) &&
		(target.barDirection === undefined || group.barDirection === target.barDirection) &&
		(target.grouping === undefined || group.grouping === target.grouping)
	)
		return structuredClone(space);
	// Office does not combine 3-D or bubble types with others: Excel's object model turns the whole
	// chart into the new type when one series takes one, so those stay whole-chart switches.
	if (
		group.is3D ||
		target.is3D ||
		!COMBO_KINDS.includes(group.kind) ||
		!COMBO_KINDS.includes(target.kind)
	)
		throw new RangeError('Series combo types need 2D column, bar, line, area or scatter families');
	const toHorizontal =
		target.kind === 'bar' && (target.barDirection ?? group.barDirection ?? 'col') === 'bar';
	const turns = horizontal(group) !== toHorizontal;
	if (turns && (group.kind === 'scatter' || target.kind === 'scatter'))
		throw new RangeError('A scatter series does not combine with horizontal bars');
	if (!axisPair(space, group))
		throw new RangeError('Combo types need a coherent category/value axis pair');
	if (group.series.length > 1 && group.extLst)
		throw new RangeError('Cannot split an unknown group extension');
	if (
		position > 0 &&
		position < group.series.length - 1 &&
		['stacked', 'percentStacked'].includes(group.grouping ?? '')
	)
		throw new RangeError('Cannot split the remaining stack into separate groups');
	// A group of the target orientation on another axis pair, which the series turns onto.
	const across = turns
		? space.plotArea.groups.find(
				(other) =>
					other !== group &&
					horizontal(other) === toHorizontal &&
					['bar', 'line', 'area'].includes(other.kind) &&
					!other.is3D &&
					axisPair(space, other) !== undefined &&
					other.axisIds.join() !== group.axisIds.join(),
			)
		: undefined;
	if (turns && !across && space.plotArea.axes.length > 2)
		throw new RangeError('A chart with secondary axes does not take another orientation here');
	return editChartSpace(space, (draft) => {
		const { group: current, groupIndex, position: at } = chartSeriesEntryAt(draft, index);
		const [x, y] = axisPair(draft, current)!;
		if (x.kind === 'val' && target.kind !== 'scatter') {
			categoryFromValueAxis(x, y);
			// Excel writes `c:smooth val="0"` on the scatter series left on the category axis;
			// PowerPoint draws one without it as smoothed lines.
			for (const other of draft.plotArea.groups)
				if (other.kind === 'scatter' && other.axisIds.join() === current.axisIds.join())
					for (const series of other.series) series.smooth ??= false;
		}
		const isolated = structuredClone(draft);
		isolated.plotArea.groups = [
			{ ...structuredClone(current), series: [structuredClone(current.series[at]!)] },
		];
		const to: ChartTypeTarget = toHorizontal ? { ...target, barDirection: 'bar' } : target;
		let converted = setChartType(isolated, to).plotArea.groups[0]!;
		if (target.kind === 'scatter' && current.kind !== 'scatter')
			markerOnlyScatter(converted.series[0]!, current.series[at]!);
		converted.axisIds = [...current.axisIds];
		if (across) converted.axisIds = [...across.axisIds];
		else if (turns) {
			const axes = secondaryPair(draft, x, toHorizontal);
			draft.plotArea.axes.push(...axes);
			converted.axisIds = axes.map((axis) => axis.id!);
		}
		// A group of the target type on the destination axes lends its settings, so the series
		// joins it rather than drawing as a second group of that type over it.
		const host = draft.plotArea.groups.find(
			(other) =>
				other !== current &&
				other.element === converted.element &&
				other.axisIds.join() === converted.axisIds.join() &&
				(target.grouping === undefined || other.grouping === target.grouping),
		);
		if (host) converted = { ...structuredClone(host), series: converted.series };
		const left = [...current.axisIds];
		const fragments = [
			{ ...structuredClone(current), series: current.series.slice(0, at) },
			converted,
			{ ...structuredClone(current), series: current.series.slice(at + 1) },
		].filter((part) => part.series.length);
		draft.plotArea.groups.splice(groupIndex, 1, ...fragments);
		// The converted group, then its neighbours of the same type.
		const groups = draft.plotArea.groups;
		const placed = groupIndex + (at > 0 ? 1 : 0);
		const join = (first: number) => {
			const [a, b] = [groups[first], groups[first + 1]];
			if (!a || !b || !sameGroupSettings(a, b)) return;
			a.series.push(...b.series);
			groups.splice(first + 1, 1);
		};
		join(placed);
		if (placed > 0) join(placed - 1);
		// A pair nothing plots on any more goes, as Office drops it.
		if (!groups.some((other) => other.axisIds.join() === left.join()))
			draft.plotArea.axes = draft.plotArea.axes.filter(
				(axis) => axis.id === undefined || !left.includes(axis.id),
			);
	});
}
