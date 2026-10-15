// Scatter and bubble plots: two numeric axes, then one dot per point (a bubble's radius follows its
// size, scaled as PowerPoint scales it: `../bubble-radius`).
import { computeBubbleRadius, maxBubbleSize, medianBubbleRadius } from '../bubble-radius';
import { formatAxisValue, type AxisScale } from './chart-scale';
import { line, n, text, textWidth, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartStroke, chartTextAttributes } from './chart-svg-appearance';

export function scatter(model: ChartViewModel, area: Rect, out: string[]): void {
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
	const bubble = model.type === 'bubble' ? model.bubble : undefined;
	// One size scale for the whole chart, so bubbles stay comparable across series.
	const biggest = maxBubbleSize(
		model.series.flatMap((s) => (s.sizes ?? []).filter((v): v is number => v !== null)),
	);
	const median = medianBubbleRadius(plot.w, plot.h);
	for (const [si, s] of model.series.entries())
		s.values.forEach((v, i) => {
			const x = s.xValues?.[i] ?? null;
			if (v === null || x === null) return;
			const filter = s.shadowFilter ? ` filter="${s.shadowFilter}"` : '';
			const fill = s.pointColors?.[i] ?? s.color;
			if (!bubble) {
				out.push(
					`<circle cx="${n(px(x, xs))}" cy="${n(py(v, ys))}" r="3" fill="${fill}"${filter}/>`,
				);
				return;
			}
			const size = s.sizes?.[i] ?? undefined;
			// A negative-size bubble is not plotted at all unless the chart shows them.
			if (size !== undefined && size < 0 && !bubble.showNegativeBubbles) return;
			const r = computeBubbleRadius(size, biggest, median, bubble);
			out.push(
				`<circle cx="${n(px(x, xs))}" cy="${n(py(v, ys))}" r="${n(r)}" fill="${fill}" fill-opacity="0.6" data-chart-series="${si}" data-chart-point="${i}"${filter}/>`,
			);
		});
}
