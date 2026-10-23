/**
 * Chart Design > Quick Layout: Layout 1-11 of a column / bar chart, each a set of chart elements
 * switched on or off together. Shared by PowerPoint, Excel, Word and Visio (they apply the same
 * catalogue; `chart/edit/quick-layout` applies one to a `ChartSpace`).
 *
 * Ground truth (PowerPoint, measured; Excel's Chart Design offers the same eleven layouts for a
 * column chart): `viewers/pptx/scripts/capture-data-galleries-com.ps1` calls
 * `Chart.ApplyLayout(1..11)` on a 4-series clustered column chart and reads back HasTitle,
 * HasLegend + Legend.Position, HasDataTable (+ ShowLegendKey), series 1's data labels (ShowValue,
 * Position = xlLabelPositionOutsideEnd), each axis' HasTitle / HasMajorGridlines /
 * HasMinorGridlines, HasAxis and the chart group's GapWidth / Overlap. The table below is that
 * read-back. Line and area charts take the same element set (their own layout lists were not
 * measured).
 *
 * @module chart/quick-layouts
 */

/** The elements a Quick Layout switches on, as its schematic tile draws them. */
export interface ChartLayoutSketch {
	title: boolean;
	legend: 'r' | 't' | 'b' | 'l' | null;
	dataLabels: boolean;
	catAxisTitle: boolean;
	valAxisTitle: boolean;
	majorGridlines: boolean;
	minorGridlines: boolean;
	valueAxis: boolean;
	dataTable: boolean;
}

export interface ChartQuickLayout extends ChartLayoutSketch {
	/** 1-based `Chart.ApplyLayout` number. */
	n: number;
	/** `c:barChart/c:gapWidth` / `c:overlap` Office sets for bar/column charts. */
	gapWidth: number;
	overlap: number;
}

const L = (
	n: number,
	flags: Partial<ChartLayoutSketch>,
	gapWidth: number,
	overlap: number,
): ChartQuickLayout => ({
	n,
	title: false,
	legend: null,
	dataLabels: false,
	catAxisTitle: false,
	valAxisTitle: false,
	majorGridlines: false,
	minorGridlines: false,
	valueAxis: true,
	dataTable: false,
	...flags,
	gapWidth,
	overlap,
});

export const CHART_QUICK_LAYOUTS: readonly ChartQuickLayout[] = [
	L(1, { title: true, legend: 'r', majorGridlines: true }, 150, 0),
	L(2, { title: true, legend: 't', dataLabels: true, valueAxis: false }, 150, -25),
	L(3, { title: true, legend: 'b', majorGridlines: true }, 75, -25),
	L(4, { legend: 'b', dataLabels: true }, 75, 0),
	L(5, { title: true, dataTable: true, valAxisTitle: true, majorGridlines: true }, 150, 0),
	L(6, { title: true, valAxisTitle: true, majorGridlines: true }, 150, 0),
	L(
		7,
		{
			legend: 'r',
			catAxisTitle: true,
			valAxisTitle: true,
			majorGridlines: true,
			minorGridlines: true,
		},
		300,
		0,
	),
	L(8, { title: true, catAxisTitle: true, valAxisTitle: true }, 0, 0),
	L(
		9,
		{ title: true, legend: 'r', catAxisTitle: true, valAxisTitle: true, majorGridlines: true },
		150,
		0,
	),
	L(10, { title: true, legend: 'r', majorGridlines: true }, 75, 40),
	L(11, { legend: 'r', majorGridlines: true }, 150, 0),
];

/** Layout `n` (1..11), or `undefined`. */
export function findChartQuickLayout(n: number): ChartQuickLayout | undefined {
	return CHART_QUICK_LAYOUTS.find((layout) => layout.n === n);
}

/** The default text Office gives an axis title a Quick Layout adds. */
export const CHART_QUICK_LAYOUT_AXIS_TITLE = 'Axis Title';
