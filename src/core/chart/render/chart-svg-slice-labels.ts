import { placeBestFitLabel } from '../pie-best-fit';
import { applyLabelManualLayout } from '../manual-layout';
import { nudgeOutsideLabels } from '../pie-label-collision';
import { measure, draw, type Measured } from './chart-svg-labels';
import { strokeAttributes } from './chart-svg-analysis';
import { n } from './chart-svg-util';
import type { ChartSeriesView, ChartViewModel } from './chart-view';
import type { ChartLabelView } from './chart-view-labels';
const GAP = 3;
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
	const items: {
		mark: SliceMark;
		label: ChartLabelView;
		m: Measured;
		x: number;
		y: number;
		outside: boolean;
	}[] = [];
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
		const position = label.position ?? (model.type === 'doughnut' ? 'ctr' : 'bestFit');
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
		items.push({
			mark,
			label,
			m,
			x,
			y,
			outside: Math.hypot(x - mark.cx, y - mark.cy) > mark.outer,
		});
	}
	const first = marks[0];
	if (!first) return;
	const nudges = nudgeOutsideLabels(
		items.flatMap((item, id) =>
			item.outside && !item.label.layout
				? [{ id, x: item.x, y: item.y, w: item.m.w, h: item.m.h }]
				: [],
		),
		{ cx: first.cx, cy: first.cy, r: first.outer },
		first.cy - first.outer,
		first.cy + first.outer,
	);
	for (const [id, item] of items.entries()) {
		const { mark, label, m } = item;
		const nudge = nudges.find((n) => n.id === id);
		let point = { x: item.x + (nudge?.dx ?? 0), y: item.y + (nudge?.dy ?? 0) };
		if (model.frameSize) point = applyLabelManualLayout(label.layout, model.frameSize, point);
		const moved = point.x !== item.x || point.y !== item.y;
		if (label.showLeaderLines !== false && (item.outside || moved)) {
			const mid = (mark.start + mark.end) / 2;
			const stroke = label.leaderLine ?? { color: '#94a3b8', width: 0.75 };
			out.push(
				`<line x1="${n(mark.cx + mark.outer * Math.sin(mid))}" y1="${n(mark.cy - mark.outer * Math.cos(mid))}" x2="${n(point.x)}" y2="${n(point.y)}" ${strokeAttributes(stroke)} data-chart-part="leaderLine"/>`,
			);
		}
		const { layout: _layout, ...fixed } = label;
		draw(m, fixed, point.x, point.y, 'middle', out);
	}
}
