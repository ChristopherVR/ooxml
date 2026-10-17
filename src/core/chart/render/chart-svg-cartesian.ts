import { formatAxisValue } from './chart-scale';
import { fit, line, n, text, textWidth, type Rect } from './chart-svg-util';
import { categoryTotals, type ChartViewModel } from './chart-view';
import { chartAreaRect, chartStroke, chartTextAttributes } from './chart-svg-appearance';
import { pointLabels, type PointMark } from './chart-svg-labels';

export interface Frame {
	plot: Rect;
	/** Value to pixel along the value axis. */
	value: (v: number) => number;
	/** Category index to the centre of its band along the category axis. */
	band: (i: number) => number;
	bandSize: number;
	/** Value to pixel along the secondary value axis of a combination chart. */
	secondary?: (v: number) => number;
}

/** Axes, gridlines and labels; returns the plot frame for the series painters. */
export function cartesianAxes(model: ChartViewModel, area: Rect, out: string[]): Frame | undefined {
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
	const second = horizontal ? undefined : model.secondaryAxis;
	const secondLabels = second?.ticks.map((t) => formatAxisValue(t, second.percent)) ?? [];
	const secondW = second
		? valueLabels
			? Math.max(...secondLabels.map((l) => textWidth(l, valueSize)), 8) + 8
			: 4
		: 0;
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
				w: Math.max(1, area.w - labelW - 4 - secondW),
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
	const frame: Frame = { plot, value, band, bandSize };
	if (second) {
		const secondSpan = second.max - second.min || 1;
		frame.secondary = (v) => plot.y + plot.h - ((v - second.min) / secondSpan) * plot.h;
		const secondAxis = chartStroke(model, 'valueAxis');
		out.push(
			line(
				plot.x + plot.w,
				plot.y,
				plot.x + plot.w,
				plot.y + plot.h,
				secondAxis.color,
				secondAxis.width,
			),
		);
		if (valueLabels)
			second.ticks.forEach((t, i) =>
				out.push(
					text(plot.x + plot.w + 4, frame.secondary!(t) + valueSize * 0.3, secondLabels[i] ?? '', {
						anchor: 'start',
						...valueAttrs,
					}),
				),
			);
	}
	return frame;
}

/** Values to plot per series, after percent-stacking. */
export function plotted(model: ChartViewModel): (number | null)[][] {
	if (model.grouping !== 'percentStacked') return model.series.map((s) => s.values);
	const totals = categoryTotals(model);
	return model.series.map((s) =>
		s.values.map((v, i) => (v === null ? null : v / (totals[i] || 1))),
	);
}

export function linesAndAreas(model: ChartViewModel, frame: Frame, out: string[]): void {
	const values = plotted(model);
	const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
	const acc = new Array<number>(model.categories.length).fill(0);
	const area = model.type === 'area';
	const marks: PointMark[] = [];
	const layers: {
		points: [number, number][];
		lower: [number, number][];
		color: string;
		filter: string;
	}[] = [];
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
			marks.push({
				series: s,
				index: i,
				x: frame.band(i),
				y: frame.value(top),
				radius: area ? 0 : 3,
			});
			lower.push([frame.band(i), frame.value(below)]);
		});
		layers.push({
			points,
			lower,
			color: s.color,
			filter: s.shadowFilter ? ` filter="${s.shadowFilter}"` : '',
		});
	});
	for (const { points, lower, color, filter } of area ? [...layers].reverse() : layers) {
		if (area) {
			const poly = [...points, ...[...lower].reverse()]
				.map(([x, y]) => `${n(x)},${n(y)}`)
				.join(' ');
			out.push(`<polygon points="${poly}" fill="${color}" fill-opacity="0.85"${filter}/>`);
			continue;
		}
		let segment: string[] = [];
		const flush = (): void => {
			if (segment.length > 1)
				out.push(
					`<polyline points="${segment.join(' ')}" fill="none" stroke="${color}" stroke-width="2.25" stroke-linejoin="round"${filter}/>`,
				);
			segment = [];
		};
		for (const [x, y] of points) {
			if (Number.isNaN(x)) flush();
			else segment.push(`${n(x)},${n(y)}`);
		}
		flush();
		for (const [x, y] of points)
			if (!Number.isNaN(x))
				out.push(`<circle cx="${n(x)}" cy="${n(y)}" r="2.5" fill="${color}"${filter}/>`);
	}
	pointLabels(model, marks, out);
}
