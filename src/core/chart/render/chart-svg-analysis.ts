// Trendlines and error bars over the series of a cartesian or scatter plot. Trendlines are fitted
// in data space (`../trendline-fit`) against the category positions 1..n, or a scatter series' x
// values, and drawn as a polyline; error bars are sized by `../error-bar-amount` and drawn with end
// caps. Stacked groups get neither (Office offers neither on a stacked chart).
import { chartErrorBarAmount } from '../error-bar-amount';
import { fitChartTrendline } from '../trendline-fit';
import { barLayout } from './chart-svg-bars';
import type { Frame } from './chart-svg-cartesian';
import { esc, n } from './chart-svg-util';
import type { ChartSeriesView, ChartViewModel } from './chart-view';
import type { ChartErrorBarsView, ChartStrokeView } from './chart-view-analysis';

/** DrawingML preset dashes (`a:prstDash`) as multiples of the line width. */
const DASHES: Readonly<Record<string, readonly number[]>> = {
	dot: [1, 3],
	dash: [4, 3],
	lgDash: [8, 3],
	dashDot: [4, 3, 1, 3],
	lgDashDot: [8, 3, 1, 3],
	lgDashDotDot: [8, 3, 1, 3, 1, 3],
	sysDash: [3, 1],
	sysDot: [1, 1],
	sysDashDot: [3, 1, 1, 1],
	sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

const CAP = 4;

function strokeAttributes(stroke: ChartStrokeView): string {
	const pattern = stroke.dash ? DASHES[stroke.dash] : undefined;
	const dash = pattern
		? ` stroke-dasharray="${pattern.map((step) => n(step * stroke.width)).join(' ')}"`
		: '';
	return `stroke="${esc(stroke.color)}" stroke-width="${n(stroke.width)}"${dash}`;
}

const segment = (x1: number, y1: number, x2: number, y2: number, stroke: string): string =>
	`<line x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" ${stroke}/>`;

/** Maps a data-space point (x along the categories or x axis, y along the values) to pixels. */
type Plane = (x: number, y: number) => [number, number];

function trendlines(
	series: ChartSeriesView,
	si: number,
	xs: number[],
	ys: number[],
	plane: Plane,
	out: string[],
): void {
	for (const trendline of series.trendlines ?? []) {
		const fit = fitChartTrendline(trendline.fit, xs, ys);
		if (fit.points.length < 2) continue;
		const points = fit.points.map(({ x, y }) => plane(x, y).map(n).join(',')).join(' ');
		out.push(
			`<polyline points="${points}" fill="none" ${strokeAttributes(trendline.stroke)} data-chart-part="trendline" data-chart-series="${si}"/>`,
		);
	}
}

/**
 * One error bar on each requested side of a point at (`x`, `y`); `along` is the axis the bar runs
 * on (`y`: the value axis, `x`: a scatter chart's x axis).
 */
function errorBar(
	bars: ChartErrorBarsView,
	at: { x: number; y: number; displayIndex: number; sourceIndex: number },
	values: readonly number[],
	plane: Plane,
	out: string[],
): void {
	const stroke = strokeAttributes(bars.stroke);
	for (const side of ['plus', 'minus'] as const) {
		if (bars.barType !== side && bars.barType !== 'both') continue;
		const amount = chartErrorBarAmount(bars.amount, values, at.displayIndex, at.sourceIndex, side);
		const sign = side === 'plus' ? 1 : -1;
		const [x1, y1] = plane(at.x, at.y);
		const [x2, y2] =
			bars.direction === 'x'
				? plane(at.x + sign * amount, at.y)
				: plane(at.x, at.y + sign * amount);
		out.push(segment(x1, y1, x2, y2, stroke));
		if (!bars.endCap) continue;
		// The cap is perpendicular to the bar.
		if (x1 === x2) out.push(segment(x2 - CAP, y2, x2 + CAP, y2, stroke));
		else out.push(segment(x2, y2 - CAP, x2, y2 + CAP, stroke));
	}
}

/** The points of a series that are plotted: category position (1-based) or x, and value. */
function plottedPoints(
	series: ChartSeriesView,
	xOf: (index: number) => number | null,
): { x: number; y: number; index: number }[] {
	return series.values.flatMap((y, index) => {
		const x = xOf(index);
		return y === null || x === null ? [] : [{ x, y, index }];
	});
}

function seriesAnalysis(
	series: ChartSeriesView,
	si: number,
	xOf: (index: number) => number | null,
	plane: Plane,
	barPlane: Plane,
	categoryAxis: boolean,
	out: string[],
): void {
	const points = plottedPoints(series, xOf);
	const xs = points.map((p) => p.x);
	const ys = points.map((p) => p.y);
	trendlines(series, series.sourceIndex ?? si, xs, ys, plane, out);
	for (const bars of series.errorBars ?? []) {
		// A category chart has no x error: its x axis is not a value axis.
		if (categoryAxis && bars.direction === 'x') continue;
		const values = bars.direction === 'x' ? xs : ys;
		points.forEach((p, displayIndex) =>
			errorBar(bars, { x: p.x, y: p.y, displayIndex, sourceIndex: p.index }, values, barPlane, out),
		);
	}
}

/** Trendlines and error bars of the first group of a bar, column, line or area chart. */
export function categoryAnalysis(model: ChartViewModel, frame: Frame, out: string[]): void {
	if (model.grouping === 'stacked' || model.grouping === 'percentStacked') return;
	const plane: Plane = (x, y) =>
		model.horizontal ? [frame.value(y), frame.band(x - 1)] : [frame.band(x - 1), frame.value(y)];
	const barred = model.type === 'bar' || model.type === 'column';
	const layout = barred ? barLayout(model, frame) : undefined;
	model.series.forEach((series, si) => {
		if (!series.trendlines && !series.errorBars) return;
		// Error bars sit on the centre of the series' own bar within its category's cluster.
		const shift = layout ? layout.offset(si) + layout.barSize / 2 : 0;
		const barPlane: Plane = (x, y) => {
			const [px, py] = plane(x, y);
			return model.horizontal ? [px, py + shift] : [px + shift, py];
		};
		seriesAnalysis(series, si, (i) => i + 1, plane, barPlane, true, out);
	});
}

/** Trendlines and error bars of a scatter or bubble chart, on its own x and y scales. */
export function scatterAnalysis(
	model: ChartViewModel,
	px: (x: number) => number,
	py: (y: number) => number,
	out: string[],
): void {
	const plane: Plane = (x, y) => [px(x), py(y)];
	model.series.forEach((series, si) => {
		if (!series.trendlines && !series.errorBars) return;
		seriesAnalysis(series, si, (i) => series.xValues?.[i] ?? null, plane, plane, false, out);
	});
}
