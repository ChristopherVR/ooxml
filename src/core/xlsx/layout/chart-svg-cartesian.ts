import { formatAxisValue, type AxisScale } from './chart-scale';
import { fit, line, n, rect, text, textWidth, type Rect } from './chart-svg-util';
import { categoryTotals, type ChartViewModel } from './chart-view';
import { chartAreaRect, chartStroke, chartTextAttributes } from './chart-svg-appearance';

/** Gap between bar clusters as a fraction of one bar (Excel's default `gapWidth` 150%). */
const GAP = 1.5;

interface Frame {
	plot: Rect;
	/** Value to pixel along the value axis. */
	value: (v: number) => number;
	/** Category index to the centre of its band along the category axis. */
	band: (i: number) => number;
	bandSize: number;
}

/** Axes, gridlines and labels; returns the plot frame for the series painters. */
function axes(model: ChartViewModel, area: Rect, out: string[]): Frame | undefined {
	const scale = model.valueAxis;
	if (!scale) return undefined;
	const labels = scale.ticks.map((t) => formatAxisValue(t, scale.percent));
	const valueAttrs = chartTextAttributes(model, 'valueAxis')!;
	const categoryAttrs = chartTextAttributes(model, 'categoryAxis')!;
	const valueSize = valueAttrs.size ?? 10;
	const categorySize = categoryAttrs.size ?? 10;
	const grid = chartStroke(model, 'gridlineMajor');
	const axis = chartStroke(model, 'categoryAxis');
	const valueLabels = model.appearance?.valueAxis?.labelsVisible !== false;
	const categoryLabels = model.appearance?.categoryAxis?.labelsVisible !== false;
	const labelW = valueLabels ? Math.max(...labels.map((l) => textWidth(l, valueSize)), 8) + 6 : 4;
	const horizontal = model.horizontal;
	const cats = Math.max(1, model.categories.length);
	const catW = Math.min(
		area.w * 0.3,
		Math.max(...model.categories.map((c) => textWidth(c, categorySize)), 8) + 6,
	);
	const plot: Rect = horizontal
		? {
				x: area.x + catW,
				y: area.y,
				w: Math.max(1, area.w - catW - 4),
				h: Math.max(1, area.h - Math.max(16, valueSize * 1.5)),
			}
		: {
				x: area.x + labelW,
				y: area.y + 4,
				w: Math.max(1, area.w - labelW - 4),
				h: Math.max(1, area.h - Math.max(20, categorySize * 1.5)),
			};
	const span = scale.max - scale.min || 1;
	const value = horizontal
		? (v: number): number => plot.x + ((v - scale.min) / span) * plot.w
		: (v: number): number => plot.y + plot.h - ((v - scale.min) / span) * plot.h;
	const bandSize = (horizontal ? plot.h : plot.w) / cats;
	const band = horizontal
		? (i: number): number => plot.y + plot.h - bandSize * (i + 0.5) // first category at the bottom, as Excel
		: (i: number): number => plot.x + bandSize * (i + 0.5);
	out.push(chartAreaRect(model, 'plotArea', plot));

	scale.ticks.forEach((t, i) => {
		const p = value(t);
		const label = labels[i] ?? '';
		if (horizontal) {
			out.push(line(p, plot.y, p, plot.y + plot.h, grid.color, grid.width));
			if (valueLabels)
				out.push(
					text(p, plot.y + plot.h + valueSize + 2, label, { anchor: 'middle', ...valueAttrs }),
				);
		} else {
			out.push(line(plot.x, p, plot.x + plot.w, p, grid.color, grid.width));
			if (valueLabels)
				out.push(text(plot.x - 4, p + valueSize * 0.3, label, { anchor: 'end', ...valueAttrs }));
		}
	});
	const zero = value(Math.min(Math.max(0, scale.min), scale.max));
	if (horizontal) out.push(line(zero, plot.y, zero, plot.y + plot.h, axis.color, axis.width));
	else out.push(line(plot.x, zero, plot.x + plot.w, zero, axis.color, axis.width));
	const step = Math.max(
		1,
		Math.ceil(cats / Math.max(1, Math.floor((horizontal ? plot.h : plot.w) / 14))),
	);
	model.categories.forEach((c, i) => {
		if (!categoryLabels || i % step !== 0) return;
		const p = band(i);
		if (horizontal)
			out.push(
				text(plot.x - 4, p + categorySize * 0.3, fit(c, catW - 6, categorySize), {
					anchor: 'end',
					...categoryAttrs,
				}),
			);
		else
			out.push(
				text(p, plot.y + plot.h + categorySize + 3, fit(c, bandSize * step, categorySize), {
					anchor: 'middle',
					...categoryAttrs,
				}),
			);
	});
	return { plot, value, band, bandSize };
}

/** Values to plot per series, after percent-stacking. */
function plotted(model: ChartViewModel): (number | null)[][] {
	if (model.grouping !== 'percentStacked') return model.series.map((s) => s.values);
	const totals = categoryTotals(model);
	return model.series.map((s) =>
		s.values.map((v, i) => (v === null ? null : v / (totals[i] || 1))),
	);
}

function bars(model: ChartViewModel, frame: Frame, out: string[]): void {
	const values = plotted(model);
	const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
	const count = Math.max(1, model.series.length);
	const barSize = frame.bandSize / ((stacked ? 1 : count) + GAP);
	const base = frame.value(0);
	const pos = new Array<number>(model.categories.length).fill(0);
	const neg = new Array<number>(model.categories.length).fill(0);
	model.series.forEach((s, si) => {
		values[si]?.forEach((v, i) => {
			if (v === null || i >= model.categories.length) return;
			let from = 0;
			let to = v;
			if (stacked) {
				const acc = v >= 0 ? pos : neg;
				from = acc[i] ?? 0;
				to = from + v;
				acc[i] = to;
			}
			const a = stacked ? frame.value(from) : base;
			const b = frame.value(to);
			const offset = stacked ? -barSize / 2 : -((count * barSize) / 2) + si * barSize;
			const center = frame.band(i);
			if (model.horizontal)
				out.push(
					rect(
						Math.min(a, b),
						center + offset,
						Math.abs(b - a),
						barSize,
						s.pointColors?.[i] ?? s.color,
					),
				);
			else
				out.push(
					rect(
						center + offset,
						Math.min(a, b),
						barSize,
						Math.abs(b - a),
						s.pointColors?.[i] ?? s.color,
					),
				);
		});
	});
}

function linesAndAreas(model: ChartViewModel, frame: Frame, out: string[]): void {
	const values = plotted(model);
	const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
	const acc = new Array<number>(model.categories.length).fill(0);
	const area = model.type === 'area';
	const layers: { points: [number, number][]; lower: [number, number][]; color: string }[] = [];
	model.series.forEach((s, si) => {
		const points: [number, number][] = [];
		const lower: [number, number][] = [];
		model.categories.forEach((_, i) => {
			const v = values[si]?.[i] ?? null;
			if (v === null && !area) {
				points.push([NaN, NaN]);
				return;
			}
			const below = stacked ? (acc[i] ?? 0) : 0;
			const top = below + (v ?? 0);
			if (stacked) acc[i] = top;
			points.push([frame.band(i), frame.value(top)]);
			lower.push([frame.band(i), frame.value(below)]);
		});
		layers.push({ points, lower, color: s.color });
	});
	for (const { points, lower, color } of area ? [...layers].reverse() : layers) {
		if (area) {
			const poly = [...points, ...[...lower].reverse()]
				.map(([x, y]) => `${n(x)},${n(y)}`)
				.join(' ');
			out.push(`<polygon points="${poly}" fill="${color}" fill-opacity="0.85"/>`);
			continue;
		}
		let segment: string[] = [];
		const flush = (): void => {
			if (segment.length > 1)
				out.push(
					`<polyline points="${segment.join(' ')}" fill="none" stroke="${color}" stroke-width="2.25" stroke-linejoin="round"/>`,
				);
			segment = [];
		};
		for (const [x, y] of points) {
			if (Number.isNaN(x)) flush();
			else segment.push(`${n(x)},${n(y)}`);
		}
		flush();
		for (const [x, y] of points)
			if (!Number.isNaN(x)) out.push(`<circle cx="${n(x)}" cy="${n(y)}" r="2.5" fill="${color}"/>`);
	}
}

function scatter(model: ChartViewModel, area: Rect, out: string[]): void {
	const ys = model.valueAxis;
	const xs = model.xAxis;
	if (!ys || !xs) return;
	const yLabels = ys.ticks.map((t) => formatAxisValue(t));
	const valueAttrs = chartTextAttributes(model, 'valueAxis')!;
	const categoryAttrs = chartTextAttributes(model, 'categoryAxis')!;
	const valueSize = valueAttrs.size ?? 10;
	const categorySize = categoryAttrs.size ?? 10;
	const grid = chartStroke(model, 'gridlineMajor');
	const xAxis = chartStroke(model, 'categoryAxis');
	const yAxis = chartStroke(model, 'valueAxis');
	const valueLabels = model.appearance?.valueAxis?.labelsVisible !== false;
	const categoryLabels = model.appearance?.categoryAxis?.labelsVisible !== false;
	const labelW = valueLabels ? Math.max(...yLabels.map((l) => textWidth(l, valueSize)), 8) + 6 : 4;
	const plot: Rect = {
		x: area.x + labelW,
		y: area.y + 4,
		w: Math.max(1, area.w - labelW - 8),
		h: Math.max(1, area.h - Math.max(20, categorySize * 1.5)),
	};
	const px = (v: number, s: AxisScale): number =>
		plot.x + ((v - s.min) / (s.max - s.min || 1)) * plot.w;
	const py = (v: number, s: AxisScale): number =>
		plot.y + plot.h - ((v - s.min) / (s.max - s.min || 1)) * plot.h;
	out.push(chartAreaRect(model, 'plotArea', plot));
	ys.ticks.forEach((t, i) => {
		out.push(line(plot.x, py(t, ys), plot.x + plot.w, py(t, ys), grid.color, grid.width));
		if (valueLabels)
			out.push(
				text(plot.x - 4, py(t, ys) + valueSize * 0.3, yLabels[i] ?? '', {
					anchor: 'end',
					...valueAttrs,
				}),
			);
	});
	for (const t of xs.ticks) {
		out.push(
			line(px(t, xs), plot.y + plot.h, px(t, xs), plot.y + plot.h + 3, xAxis.color, xAxis.width),
		);
		if (categoryLabels)
			out.push(
				text(px(t, xs), plot.y + plot.h + categorySize + 3, formatAxisValue(t), {
					anchor: 'middle',
					...categoryAttrs,
				}),
			);
	}
	out.push(
		line(plot.x, plot.y + plot.h, plot.x + plot.w, plot.y + plot.h, xAxis.color, xAxis.width),
	);
	out.push(line(plot.x, plot.y, plot.x, plot.y + plot.h, yAxis.color, yAxis.width));
	for (const s of model.series)
		s.values.forEach((v, i) => {
			const x = s.xValues?.[i] ?? null;
			if (v === null || x === null) return;
			out.push(
				`<circle cx="${n(px(x, xs))}" cy="${n(py(v, ys))}" r="3" fill="${s.pointColors?.[i] ?? s.color}"/>`,
			);
		});
}

/** Paints bar, column, line, area and scatter charts into `area`. */
export function cartesianSvg(model: ChartViewModel, area: Rect): string {
	const out: string[] = [];
	if (model.type === 'scatter') {
		scatter(model, area, out);
		return out.join('');
	}
	const frame = axes(model, area, out);
	if (!frame) return '';
	if (model.type === 'bar' || model.type === 'column') bars(model, frame, out);
	else linesAndAreas(model, frame, out);
	return out.join('');
}
