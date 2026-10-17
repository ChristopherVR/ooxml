import { axisTickText, axisGridlines, axisLine } from './chart-svg-axes';
import { chartAxisRatio, chartAxisLabels } from './chart-view-axes';
// Scatter and bubble plots: two numeric axes, then one dot per point (a bubble's radius follows its
// size, scaled as PowerPoint scales it: `../bubble-radius`).
import { computeBubbleRadius, maxBubbleSize, medianBubbleRadius } from '../bubble-radius';
import { scatterAnalysis } from './chart-svg-analysis';
import { type AxisScale } from './chart-scale';
import { line, n, text, textWidth, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { chartAreaRect, chartStroke, chartTextAttributes } from './chart-svg-appearance';
import { pointLabels, type PointMark } from './chart-svg-labels';
import { chartMarkerRadius, chartSeriesMarkerSvg } from './chart-svg-markers';

export function scatter(model: ChartViewModel, area: Rect, out: string[]): void {
	const ys = model.valueAxis;
	const xs = model.xAxis;
	if (!ys || !xs) return;
	const yLabels = ys.ticks.map((t) => axisTickText(t, model.axes?.value));
	const valueAttrs = chartTextAttributes(model, 'valueAxis')!;
	const categoryAttrs = chartTextAttributes(model, 'categoryAxis')!;
	const valueSize = valueAttrs.size ?? 10;
	const categorySize = categoryAttrs.size ?? 10;
	const xAxis = chartStroke(model, 'categoryAxis');
	const valueLabels =
		chartAxisLabels(model.axes?.value) && model.appearance?.valueAxis?.labelsVisible !== false;
	const categoryLabels =
		chartAxisLabels(model.axes?.category) &&
		model.appearance?.categoryAxis?.labelsVisible !== false;
	const labelW = valueLabels ? Math.max(...yLabels.map((l) => textWidth(l, valueSize)), 8) + 6 : 4;
	const plot: Rect = {
		x: area.x + labelW,
		y: area.y + 4,
		w: Math.max(1, area.w - labelW - 8),
		h: Math.max(1, area.h - Math.max(20, categorySize * 1.5)),
	};
	const bubble = model.type === 'bubble' ? model.bubble : undefined;
	// One size scale for the whole chart, so bubbles stay comparable across series.
	const biggest = maxBubbleSize(
		model.series.flatMap((s) => (s.sizes ?? []).filter((v): v is number => v !== null)),
	);
	const median = medianBubbleRadius(plot.w, plot.h);
	const radiusOf = (size: number | undefined): number | undefined =>
		bubble && !(size !== undefined && size < 0 && !bubble.showNegativeBubbles)
			? computeBubbleRadius(size, biggest, median, bubble)
			: undefined;
	// Like Excel, both axes extend by the largest radius so an edge bubble is never clipped.
	let reach = 0;
	if (bubble)
		for (const s of model.series)
			s.values.forEach((_, i) => {
				reach = Math.max(reach, radiusOf(s.sizes?.[i] ?? undefined) ?? 0);
			});
	const widen = (s: AxisScale, length: number): AxisScale => {
		const pad = (reach * (s.max - s.min || 1)) / Math.max(1, length - 2 * reach);
		return { ...s, min: s.min - pad, max: s.max + pad };
	};
	const xView = model.axes?.category ? xs : widen(xs, plot.w);
	const yView = model.axes?.value ? ys : widen(ys, plot.h);
	const px = (v: number): number =>
		plot.x + chartAxisRatio(v, xView, model.axes?.category) * plot.w;
	const py = (v: number): number =>
		plot.y + plot.h - chartAxisRatio(v, yView, model.axes?.value) * plot.h;
	out.push(chartAreaRect(model, 'plotArea', plot));
	out.push(axisGridlines(model, model.axes?.value, ys.ticks, py, plot, false));
	if (model.axes?.category)
		out.push(axisGridlines(model, model.axes.category, xs.ticks, px, plot, true));
	ys.ticks.forEach((t, i) => {
		if (valueLabels)
			out.push(
				text(plot.x - 4, py(t) + valueSize * 0.3, yLabels[i] ?? '', {
					anchor: 'end',
					...valueAttrs,
				}),
			);
	});
	for (const t of xs.ticks) {
		if (!model.axes?.category?.source.deleted)
			out.push(line(px(t), plot.y + plot.h, px(t), plot.y + plot.h + 3, xAxis.color, xAxis.width));
		if (categoryLabels)
			out.push(
				text(px(t), plot.y + plot.h + categorySize + 3, axisTickText(t, model.axes?.category), {
					anchor: 'middle',
					...categoryAttrs,
				}),
			);
	}
	out.push(
		axisLine(
			model,
			model.axes?.category,
			'categoryAxis',
			plot.x,
			plot.y + plot.h,
			plot.x + plot.w,
			plot.y + plot.h,
		),
	);
	out.push(
		axisLine(model, model.axes?.value, 'valueAxis', plot.x, plot.y, plot.x, plot.y + plot.h),
	);
	const marks: PointMark[] = [];
	for (const [si, s] of model.series.entries())
		s.values.forEach((v, i) => {
			const x = s.xValues?.[i] ?? null;
			if (v === null || x === null) return;
			const filter = s.shadowFilter ? ` filter="${s.shadowFilter}"` : '';
			const fill = s.pointColors?.[i] ?? s.color;
			if (!bubble) {
				out.push(chartSeriesMarkerSvg(s, i, px(x), py(v)));
				marks.push({ series: s, index: i, x: px(x), y: py(v), radius: chartMarkerRadius(s, i) });
				return;
			}
			// A negative-size bubble is not plotted at all unless the chart shows them.
			const r = radiusOf(s.sizes?.[i] ?? undefined);
			if (r === undefined) return;
			marks.push({ series: s, index: i, x: px(x), y: py(v), radius: r });
			out.push(
				`<circle cx="${n(px(x))}" cy="${n(py(v))}" r="${n(r)}" fill="${fill}" fill-opacity="0.6" data-chart-series="${si}" data-chart-point="${i}"${filter}/>`,
			);
		});
	pointLabels(model, marks, out);
	scatterAnalysis(model, px, py, out);
}
