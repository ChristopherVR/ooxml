import { formatAxisValue } from './chart-scale';
import { fit, line, n, text, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartStroke, chartTextAttributes } from './chart-svg-appearance';

const polar = (cx: number, cy: number, r: number, angle: number): [number, number] => [
	cx + r * Math.sin(angle),
	cy - r * Math.cos(angle),
];

/** An annular (or full) slice path from angle `a0` to `a1`, clockwise from 12 o'clock. */
function slicePath(
	cx: number,
	cy: number,
	r: number,
	inner: number,
	a0: number,
	a1: number,
): string {
	const sweep = a1 - a0;
	if (sweep >= Math.PI * 2 - 1e-9) {
		// A full circle cannot be one arc; draw two halves.
		const mid = a0 + Math.PI;
		return slicePath(cx, cy, r, inner, a0, mid) + slicePath(cx, cy, r, inner, mid, a1);
	}
	const large = sweep > Math.PI ? 1 : 0;
	const [x0, y0] = polar(cx, cy, r, a0);
	const [x1, y1] = polar(cx, cy, r, a1);
	if (inner <= 0)
		return `M${n(cx)},${n(cy)} L${n(x0)},${n(y0)} A${n(r)},${n(r)} 0 ${large} 1 ${n(x1)},${n(y1)} Z`;
	const [ix1, iy1] = polar(cx, cy, inner, a1);
	const [ix0, iy0] = polar(cx, cy, inner, a0);
	return `M${n(x0)},${n(y0)} A${n(r)},${n(r)} 0 ${large} 1 ${n(x1)},${n(y1)} L${n(ix1)},${n(iy1)} A${n(inner)},${n(inner)} 0 ${large} 0 ${n(ix0)},${n(iy0)} Z`;
}

/** Pie (first series) and doughnut (one ring per series) charts. */
export function pieSvg(model: ChartViewModel, area: Rect): string {
	const out: string[] = [chartAreaRect(model, 'plotArea', area)];
	const cx = area.x + area.w / 2;
	const cy = area.y + area.h / 2;
	const radius = Math.max(1, Math.min(area.w, area.h) / 2 - 4);
	const rings = model.type === 'doughnut' ? model.series : model.series.slice(0, 1);
	const hole = model.type === 'doughnut' ? radius * 0.5 : 0;
	const ringSize = (radius - hole) / Math.max(1, rings.length);
	rings.forEach((s, ri) => {
		const outer = radius - ri * ringSize;
		const inner = model.type === 'doughnut' ? outer - ringSize : 0;
		const total = s.values.reduce<number>((sum, v) => sum + (v !== null && v > 0 ? v : 0), 0);
		if (total <= 0) return;
		let angle = 0;
		s.values.forEach((v, i) => {
			if (v === null || v <= 0) return;
			const sweep = (v / total) * Math.PI * 2;
			const color = s.pointColors?.[i] ?? s.color;
			out.push(
				`<path d="${slicePath(cx, cy, outer, inner, angle, angle + sweep)}" fill="${color}" stroke="#FFFFFF" stroke-width="1"/>`,
			);
			angle += sweep;
		});
	});
	return out.join('');
}

/** Radar charts: one closed polyline per series over category spokes. */
export function radarSvg(model: ChartViewModel, area: Rect): string {
	const out: string[] = [chartAreaRect(model, 'plotArea', area)];
	const scale = model.valueAxis;
	const count = model.categories.length;
	if (!scale || count < 1) return '';
	const valueAttrs = chartTextAttributes(model, 'valueAxis', 9);
	const categoryAttrs = chartTextAttributes(model, 'categoryAxis');
	const valueSize = valueAttrs.size ?? 9;
	const categorySize = categoryAttrs.size ?? 10;
	const grid = chartStroke(model, 'gridlineMajor');
	const axis = chartStroke(model, 'categoryAxis');
	const cx = area.x + area.w / 2;
	const cy = area.y + area.h / 2 + 4;
	const radius = Math.max(1, Math.min(area.w, area.h) / 2 - Math.max(18, categorySize * 1.8));
	const angle = (i: number): number => (i / count) * Math.PI * 2;
	const r = (v: number): number => ((v - scale.min) / (scale.max - scale.min || 1)) * radius;
	const labelEvery = Math.max(
		1,
		Math.ceil(12 / Math.max(1, radius / Math.max(1, scale.ticks.length - 1))),
	);
	for (const [k, t] of scale.ticks.entries()) {
		const ring = Array.from({ length: count }, (_, i) => polar(cx, cy, r(t), angle(i)))
			.map(([x, y]) => `${n(x)},${n(y)}`)
			.join(' ');
		out.push(
			`<polygon points="${ring}" fill="none" stroke="${grid.color}" stroke-width="${n(grid.width)}"/>`,
		);
		if (model.appearance?.valueAxis?.labelsVisible !== false && k % labelEvery === 0)
			out.push(
				text(cx + 3, cy - r(t) + valueSize * 0.3, formatAxisValue(t, scale.percent), valueAttrs),
			);
	}
	model.categories.forEach((c, i) => {
		const [x, y] = polar(cx, cy, radius, angle(i));
		out.push(line(cx, cy, x, y, axis.color, axis.width));
		const [lx, ly] = polar(cx, cy, radius + 10, angle(i));
		const anchor = Math.abs(lx - cx) < 2 ? 'middle' : lx > cx ? 'start' : 'end';
		if (model.appearance?.categoryAxis?.labelsVisible !== false)
			out.push(
				text(lx, ly + categorySize * 0.3, fit(c, area.w / 4, categorySize), {
					anchor,
					...categoryAttrs,
				}),
			);
	});
	for (const s of model.series) {
		const points = s.values
			.slice(0, count)
			.map((v, i) => polar(cx, cy, r(v ?? scale.min), angle(i)))
			.map(([x, y]) => `${n(x)},${n(y)}`)
			.join(' ');
		out.push(`<polygon points="${points}" fill="none" stroke="${s.color}" stroke-width="2"/>`);
	}
	return out.join('');
}
