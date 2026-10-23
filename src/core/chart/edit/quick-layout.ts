// Chart Design > Quick Layout over a `ChartSpace`. Semantics: PowerPoint's `Chart.ApplyLayout`
// (`chart/quick-layouts`, measured on a column chart; Excel's layouts for a column chart are the
// same eleven): a layout switches the title, legend, data labels (value, outside end on bars), axis
// titles ("Axis Title"), value-axis gridlines, the value axis itself and the data table on or off
// together, and sets the bar group's gap width and overlap. Axes keep their other options. Line and
// area charts take the same element set, without the bar spacing; other families have no layouts.
import type { ChartSpace } from '../model';
import {
	CHART_QUICK_LAYOUT_AXIS_TITLE,
	CHART_QUICK_LAYOUTS,
	findChartQuickLayout,
	type ChartQuickLayout,
} from '../quick-layouts';
import { editChartSpace } from './draft';
import { setChartDataLabels, setChartLegend } from './legend-labels';
import { newChartTitle, setChartTitleVisible } from './title';

/** The chart group families a Quick Layout applies to. */
export const CHART_QUICK_LAYOUT_KINDS: ReadonlySet<string> = new Set(['bar', 'line', 'area']);

/** The parts of a layout; `applyChartQuickLayout` applies every part unless told otherwise. */
export type ChartQuickLayoutPart =
	| 'title'
	| 'legend'
	| 'dataLabels'
	| 'axes'
	| 'dataTable'
	| 'spacing';

const ALL_PARTS: readonly ChartQuickLayoutPart[] = [
	'title',
	'legend',
	'dataLabels',
	'axes',
	'dataTable',
	'spacing',
];

/** Whether the chart's first group takes Quick Layouts (bar, column, line, area). */
export function chartSupportsQuickLayout(space: ChartSpace): boolean {
	const group = space.plotArea.groups[0];
	return !!group && CHART_QUICK_LAYOUT_KINDS.has(group.kind);
}

function applyAxes(draft: ChartSpace, layout: ChartQuickLayout): void {
	for (const axis of draft.plotArea.axes) {
		const value = axis.kind === 'val';
		if (!value && axis.kind !== 'cat' && axis.kind !== 'date') continue;
		const titled = value ? layout.valAxisTitle : layout.catAxisTitle;
		if (!titled) delete axis.title;
		else axis.title ??= newChartTitle(CHART_QUICK_LAYOUT_AXIS_TITLE);
		axis.majorGridlines = value && layout.majorGridlines;
		axis.minorGridlines = value && layout.minorGridlines;
		axis.deleted = value ? !layout.valueAxis : false;
	}
}

function applyDataTable(draft: ChartSpace, layout: ChartQuickLayout): void {
	if (!layout.dataTable) {
		delete draft.plotArea.dataTable;
		return;
	}
	draft.plotArea.dataTable = {
		...draft.plotArea.dataTable,
		showHorizontalBorder: true,
		showVerticalBorder: true,
		showOutline: true,
		showKeys: true,
	};
}

/**
 * Applies Quick Layout `n` (1..11). `only` limits it to some parts (a product that already writes
 * the title or legend from its own model applies the rest). Throws for an unknown layout or a
 * chart family without layouts.
 */
export function applyChartQuickLayout(
	space: ChartSpace,
	n: number,
	only: readonly ChartQuickLayoutPart[] = ALL_PARTS,
): ChartSpace {
	const layout = findChartQuickLayout(n);
	if (!layout) throw new RangeError(`No chart quick layout ${n}`);
	if (!chartSupportsQuickLayout(space))
		throw new RangeError('Quick layouts apply to bar, column, line and area charts');
	const parts = new Set(only);
	let next = space;
	if (parts.has('title')) next = setChartTitleVisible(next, layout.title);
	if (parts.has('legend'))
		next = setChartLegend(
			next,
			layout.legend ? { visible: true, position: layout.legend } : { visible: false },
		);
	if (parts.has('dataLabels'))
		next = setChartDataLabels(
			next,
			layout.dataLabels ? { show: true, showValue: true } : { show: false },
		);
	return editChartSpace(next, (draft) => {
		if (parts.has('dataLabels') && layout.dataLabels)
			for (const group of draft.plotArea.groups)
				if (group.kind === 'bar' && group.dataLabels) group.dataLabels.position = 'outEnd';
		if (parts.has('axes')) applyAxes(draft, layout);
		if (parts.has('dataTable')) applyDataTable(draft, layout);
		if (parts.has('spacing'))
			for (const group of draft.plotArea.groups) {
				if (group.kind !== 'bar') continue;
				group.gapWidth = layout.gapWidth;
				group.overlap = layout.overlap;
			}
	});
}

/** Whether the chart's elements are exactly layout `layout`'s. */
function matches(space: ChartSpace, layout: ChartQuickLayout): boolean {
	const axes = space.plotArea.axes;
	const val = axes.find((axis) => axis.kind === 'val');
	const cat = axes.find((axis) => axis.kind === 'cat' || axis.kind === 'date');
	const bar = space.plotArea.groups.find((group) => group.kind === 'bar');
	const labels = space.plotArea.groups.some(
		(group) => group.dataLabels && group.dataLabels.deleted !== true && group.dataLabels.showValue,
	);
	const legend = space.legend ? (space.legend.position ?? 'r') : null;
	return (
		(!bar ||
			((bar.gapWidth ?? 150) === layout.gapWidth && (bar.overlap ?? 0) === layout.overlap)) &&
		(space.title !== undefined && space.autoTitleDeleted !== true) === layout.title &&
		legend === layout.legend &&
		labels === layout.dataLabels &&
		!!space.plotArea.dataTable === layout.dataTable &&
		!!cat?.title === layout.catAxisTitle &&
		!!val?.title === layout.valAxisTitle &&
		!!val?.majorGridlines === layout.majorGridlines &&
		!!val?.minorGridlines === layout.minorGridlines &&
		!val?.deleted === layout.valueAxis
	);
}

/** The Quick Layout the chart currently shows (its number), if its elements match one exactly. */
export function chartQuickLayoutOf(space: ChartSpace): number | undefined {
	if (!chartSupportsQuickLayout(space)) return undefined;
	return CHART_QUICK_LAYOUTS.find((layout) => matches(space, layout))?.n;
}
