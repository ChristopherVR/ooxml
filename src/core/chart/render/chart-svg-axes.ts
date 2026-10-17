import { formatChartNumber } from '../number-format';
import { chartLineDashArray } from '../line-dash';
import type { ChartAxisView, ChartAxisPaint } from './chart-view-axes';
import type { ChartViewModel } from './chart-view';
import { chartStroke } from './chart-svg-appearance';
import { formatAxisValue } from './chart-scale';
import { esc, line, type Rect } from './chart-svg-util';

/** Native number formats win over the painter's compact fallback. */
export function axisTickText(value: number, axis?: ChartAxisView, percent = false): string {
	return (
		formatChartNumber(value, axis?.source.numberFormat?.formatCode) ??
		formatAxisValue(value, percent)
	);
}

export function axisLine(
	model: ChartViewModel,
	axis: ChartAxisView | undefined,
	part: 'categoryAxis' | 'valueAxis',
	x1: number,
	y1: number,
	x2: number,
	y2: number,
): string {
	if (axis?.source.deleted) return '';
	return styledAxisLine(x1, y1, x2, y2, axis?.line, chartStroke(model, part));
}

function styledAxisLine(
	x1: number,
	y1: number,
	x2: number,
	y2: number,
	paint: ChartAxisPaint | undefined,
	fallback: { color: string; width: number },
): string {
	const color = paint?.color ?? fallback.color;
	const width = paint?.width ?? fallback.width;
	if (color === 'none' || width === 0) return '';
	const dash = chartLineDashArray(paint?.dash, width);
	return line(x1, y1, x2, y2, esc(color), width).replace(
		'/>',
		dash ? ` stroke-dasharray="${dash}"/>` : '/>',
	);
}

/** Gridline presence belongs to its axis, independently of whether its labels are hidden. */
export function axisGridlines(
	model: ChartViewModel,
	axis: ChartAxisView | undefined,
	ticks: readonly number[],
	position: (value: number) => number,
	plot: Rect,
	vertical: boolean,
): string {
	const draw = (
		values: readonly number[],
		paint: ChartAxisPaint | undefined,
		minor: boolean,
	): string => {
		const fallback = chartStroke(model, minor ? 'gridlineMinor' : 'gridlineMajor');
		return values
			.map((value) => {
				const p = position(value);
				return vertical
					? styledAxisLine(p, plot.y, p, plot.y + plot.h, paint, fallback)
					: styledAxisLine(plot.x, p, plot.x + plot.w, p, paint, fallback);
			})
			.join('');
	};
	const major = !axis || axis.source.majorGridlines ? draw(ticks, axis?.majorGrid, false) : '';
	const minor = axis?.source.minorGridlines ? draw(axis.minorTicks, axis.minorGrid, true) : '';
	return major + minor;
}
