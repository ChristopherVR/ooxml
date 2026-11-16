import type { ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { setChartType, type ChartTypeTarget } from './type';

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
 * The families one series can take in a category-axis combination. A scatter series there shares
 * the category and value axes of the group it leaves, with its categories as x values: the
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

/**
 * Changes one category-axis combo series without changing document-order selections or axis ids.
 * The series joins a neighbouring group of the type it takes when the two would differ only in
 * their series, as Office keeps one group per type and axis pair.
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
		!COMBO_KINDS.includes(target.kind) ||
		group.barDirection === 'bar' ||
		target.barDirection === 'bar'
	)
		throw new RangeError(
			'Series combo types need vertical 2D column, line, area or scatter families',
		);
	const pair = group.axisIds.map((id) => space.plotArea.axes.find((axis) => axis.id === id));
	if (
		pair.length !== 2 ||
		!pair[0] ||
		!pair[1] ||
		!['cat', 'date'].includes(pair[0].kind) ||
		pair[1].kind !== 'val' ||
		pair[0].crossAxisId !== pair[1].id ||
		pair[1].crossAxisId !== pair[0].id
	)
		throw new RangeError('Combo types need a coherent category/value axis pair');
	if (group.series.length > 1 && group.extLst)
		throw new RangeError('Cannot split an unknown group extension');
	if (
		position > 0 &&
		position < group.series.length - 1 &&
		['stacked', 'percentStacked'].includes(group.grouping ?? '')
	)
		throw new RangeError('Cannot split the remaining stack into separate groups');
	const isolated = structuredClone(space);
	isolated.plotArea.groups = [
		{ ...structuredClone(group), series: [structuredClone(group.series[position]!)] },
	];
	const converted = setChartType(isolated, target).plotArea.groups[0]!;
	if (target.kind === 'scatter' && group.kind !== 'scatter')
		markerOnlyScatter(converted.series[0]!, group.series[position]!);
	return editChartSpace(space, (draft) => {
		const { group: current, groupIndex, position: at } = chartSeriesEntryAt(draft, index);
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
	});
}
