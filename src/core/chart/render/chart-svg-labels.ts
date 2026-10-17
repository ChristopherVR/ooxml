// Paints the data labels of a chart view model over its marks. Each painter that draws marks
// (bars, lines, scatter points, pie slices, radar vertices) records where they are and hands the
// records here, so the labels land on top of every series and follow the placement Office uses for
// `c:dLblPos`.
import { placeBestFitLabel } from '../pie-best-fit';
import { chartTextAttributes } from './chart-svg-appearance';
import { text, textWidth } from './chart-svg-util';
import type { ChartSeriesView, ChartViewModel } from './chart-view';
import type { ChartLabelView } from './chart-view-labels';

/** Space between a mark and an outside label. */
const GAP = 3;
/** Line spacing of a multi-line label, as a multiple of the font size. */
const LINE = 1.2;
const NAMED_COLORS = /^(black|blue|cyan|green|magenta|red|white|yellow)$/iu;

type Anchor = 'start' | 'middle' | 'end';

interface Measured {
	lines: string[];
	size: number;
	w: number;
	h: number;
	attrs: ReturnType<typeof chartTextAttributes>;
}

function measure(model: ChartViewModel, label: ChartLabelView): Measured {
	const attrs = chartTextAttributes(model, 'dataLabel', 10);
	const size = attrs.size ?? 10;
	const lines = label.text.split('\n');
	return {
		lines,
		size,
		w: Math.max(...lines.map((l) => textWidth(l, size))),
		h: lines.length * size * LINE,
		attrs,
	};
}

/** One label with its text block centred vertically on `y`, aligned at `x` by `anchor`. */
function draw(
	m: Measured,
	label: ChartLabelView,
	x: number,
	y: number,
	anchor: Anchor,
	out: string[],
): void {
	const fill =
		label.color && NAMED_COLORS.test(label.color) ? label.color.toLowerCase() : undefined;
	m.lines.forEach((line, k) => {
		const lineY = y - m.h / 2 + (k + 0.5) * m.size * LINE + m.size * 0.35;
		out.push(text(x, lineY, line, { ...m.attrs, ...(fill ? { fill } : {}), anchor }));
	});
}

/** A bar or column as drawn: its box and the sign of its value (which end is the "end"). */
export interface BarMark {
	series: ChartSeriesView;
	index: number;
	x: number;
	y: number;
	w: number;
	h: number;
	negative: boolean;
}

/** Labels of bars and columns: `ctr`, `inEnd`, `inBase`, `outEnd`; clustered bars default to `outEnd`. */
export function barLabels(model: ChartViewModel, marks: readonly BarMark[], out: string[]): void {
	const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
	for (const mark of marks) {
		const label = mark.series.labels?.[mark.index];
		if (!label) continue;
		const m = measure(model, label);
		let position = label.position ?? (stacked ? 'ctr' : 'outEnd');
		if (stacked && position === 'outEnd') position = 'ctr';
		const cx = mark.x + mark.w / 2;
		const cy = mark.y + mark.h / 2;
		if (model.horizontal) {
			// The end of a horizontal bar is its right edge (its left edge for a negative value).
			const end = mark.negative ? mark.x : mark.x + mark.w;
			const base = mark.negative ? mark.x + mark.w : mark.x;
			const dir = mark.negative ? -1 : 1;
			const side = (at: number, outward: boolean): [number, Anchor] => [
				at + (outward ? dir : -dir) * GAP,
				(outward ? dir : -dir) > 0 ? 'start' : 'end',
			];
			const [x, anchor]: [number, Anchor] =
				position === 'ctr'
					? [cx, 'middle']
					: position === 'inEnd'
						? side(end, false)
						: position === 'inBase'
							? side(base, true)
							: side(end, true);
			draw(m, label, x, cy, anchor, out);
			continue;
		}
		const top = mark.negative ? mark.y + mark.h : mark.y;
		const bottom = mark.negative ? mark.y : mark.y + mark.h;
		const dir = mark.negative ? 1 : -1;
		const y =
			position === 'ctr'
				? cy
				: position === 'inEnd'
					? top - dir * (GAP + m.h / 2)
					: position === 'inBase'
						? bottom + dir * (GAP + m.h / 2)
						: top + dir * (GAP + m.h / 2);
		draw(m, label, cx, y, 'middle', out);
	}
}

/** A plotted point: a line or scatter vertex, a bubble (with its radius), an area crest. */
export interface PointMark {
	series: ChartSeriesView;
	index: number;
	x: number;
	y: number;
	radius: number;
}

/** Labels of points: `r` (the default), `l`, `t`, `b` and `ctr`; areas label at the point. */
export function pointLabels(
	model: ChartViewModel,
	marks: readonly PointMark[],
	out: string[],
): void {
	for (const mark of marks) {
		const label = mark.series.labels?.[mark.index];
		if (!label) continue;
		const m = measure(model, label);
		const reach = mark.radius + GAP;
		switch (model.type === 'area' ? 'ctr' : (label.position ?? 'r')) {
			case 'ctr':
				draw(m, label, mark.x, mark.y, 'middle', out);
				break;
			case 't':
				draw(m, label, mark.x, mark.y - reach - m.h / 2, 'middle', out);
				break;
			case 'b':
				draw(m, label, mark.x, mark.y + reach + m.h / 2, 'middle', out);
				break;
			case 'l':
				draw(m, label, mark.x - reach, mark.y, 'end', out);
				break;
			default:
				draw(m, label, mark.x + reach, mark.y, 'start', out);
		}
	}
}

/** A pie or doughnut slice: its ring and its edge angles (radians clockwise from 12 o'clock). */
export interface SliceMark {
	series: ChartSeriesView;
	index: number;
	cx: number;
	cy: number;
	outer: number;
	inner: number;
	start: number;
	end: number;
}

/**
 * Labels of pie slices: `ctr`, `inEnd`, `outEnd` and `bestFit` (the default), the latter with
 * Office's placement (`placeBestFitLabel`). A doughnut ring only ever centres its labels.
 */
export function sliceLabels(
	model: ChartViewModel,
	marks: readonly SliceMark[],
	out: string[],
): void {
	for (const mark of marks) {
		const label = mark.series.labels?.[mark.index];
		if (!label) continue;
		const m = measure(model, label);
		const mid = (mark.start + mark.end) / 2;
		const ux = Math.sin(mid);
		const uy = -Math.cos(mid);
		const place = (radius: number): [number, number] => [
			mark.cx + radius * ux,
			mark.cy + radius * uy,
		];
		const position = model.type === 'doughnut' ? 'ctr' : (label.position ?? 'bestFit');
		let x: number;
		let y: number;
		if (position === 'ctr') [x, y] = place(mark.inner + (mark.outer - mark.inner) / 2);
		else if (position === 'inEnd') {
			const extent = (Math.abs(ux) * m.w) / 2 + (Math.abs(uy) * m.h) / 2;
			[x, y] = place(Math.max(0, mark.outer - GAP - extent));
		} else if (position === 'outEnd') {
			const extent = (Math.abs(ux) * m.w) / 2 + (Math.abs(uy) * m.h) / 2;
			[x, y] = place(mark.outer + GAP + extent);
		} else {
			// The placement works in the SVG angle convention (0 at 3 o'clock).
			const placed = placeBestFitLabel(
				{
					cx: mark.cx,
					cy: mark.cy,
					outerR: mark.outer,
					startAngle: mark.start - Math.PI / 2,
					endAngle: mark.end - Math.PI / 2,
				},
				m.w,
				m.h,
			);
			x = placed.x;
			y = placed.y;
		}
		draw(m, label, x, y, 'middle', out);
	}
}

/** Whether any label of the model sits outside the plot (so a pie must leave room for it). */
export function hasOutsideLabels(model: ChartViewModel): boolean {
	return model.series.some((s) =>
		s.labels?.some(
			(l) => l !== undefined && ['bestFit', 'outEnd'].includes(l.position ?? 'bestFit'),
		),
	);
}
