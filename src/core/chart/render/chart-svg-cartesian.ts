import { chartMarkerRadius, chartSeriesMarkerSvg } from './chart-svg-markers';
import { axisTickText, axisGridlines, axisLine } from './chart-svg-axes';
import { chartAxisRatio, chartAxisLabels } from './chart-view-axes';
import { fit, n, text, textWidth, type Rect } from './chart-svg-util';
import { categoryTotals, type ChartViewModel } from './chart-view';
import { chartAreaRect, chartTextAttributes } from './chart-svg-appearance';
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
	const labels = scale.ticks.map((t) => axisTickText(t, model.axes?.value, scale.percent));
	const valueAttrs = chartTextAttributes(model, 'valueAxis')!;
	const categoryAttrs = chartTextAttributes(model, 'categoryAxis')!;
	const valueSize = valueAttrs.size ?? 10;
	const categorySize = categoryAttrs.size ?? 10;
	const valueLabels =
		chartAxisLabels(model.axes?.value) && model.appearance?.valueAxis?.labelsVisible !== false;
	const categoryLabels =
		chartAxisLabels(model.axes?.category) &&
		model.appearance?.categoryAxis?.labelsVisible !== false;
	const labelW = valueLabels ? Math.max(...labels.map((l) => textWidth(l, valueSize)), 8) + 6 : 4;
	const horizontal = model.horizontal;
	const cats = Math.max(1, model.categories.length);
	const catW = Math.min(
		area.w * 0.3,
		Math.max(...model.categories.map((c) => textWidth(c, categorySize)), 8) + 6,
	);
	const second = horizontal ? undefined : model.secondaryAxis;
	const secondLabels =
		second?.ticks.map((t) => axisTickText(t, model.axes?.secondary, second.percent)) ?? [];
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
	const value = horizontal
		? (v: number): number => plot.x + chartAxisRatio(v, scale, model.axes?.value) * plot.w
		: (v: number): number => plot.y + plot.h - chartAxisRatio(v, scale, model.axes?.value) * plot.h;
	const bandSize = (horizontal ? plot.h : plot.w) / cats;
	const index = (i: number): number =>
		model.axes?.category?.range.reverseOrder ? cats - 1 - i : i;
	const band = horizontal
		? (i: number): number => plot.y + plot.h - bandSize * (index(i) + 0.5) // first category at the bottom, as Excel
		: (i: number): number => plot.x + bandSize * (index(i) + 0.5);
	out.push(chartAreaRect(model, 'plotArea', plot));
	out.push(axisGridlines(model, model.axes?.value, scale.ticks, value, plot, horizontal));
	if (model.axes?.category?.source.majorGridlines)
		out.push(
			axisGridlines(
				model,
				model.axes.category,
				model.categories.map((_, i) => i),
				band,
				plot,
				!horizontal,
			),
		);

	scale.ticks.forEach((t, i) => {
		const p = value(t);
		const label = labels[i] ?? '';
		if (horizontal) {
			if (valueLabels)
				out.push(
					text(p, plot.y + plot.h + valueSize + 2, label, { anchor: 'middle', ...valueAttrs }),
				);
		} else {
			if (valueLabels)
				out.push(text(plot.x - 4, p + valueSize * 0.3, label, { anchor: 'end', ...valueAttrs }));
		}
	});
	const source = model.axes?.value?.source;
	const crossing =
		source?.crosses === 'max'
			? scale.max
			: source?.crosses === 'min'
				? scale.min
				: (source?.crossesAt ?? 0);
	const zero = value(Math.min(Math.max(crossing, scale.min), scale.max));
	if (horizontal)
		out.push(
			axisLine(model, model.axes?.category, 'categoryAxis', zero, plot.y, zero, plot.y + plot.h),
		);
	else
		out.push(
			axisLine(model, model.axes?.category, 'categoryAxis', plot.x, zero, plot.x + plot.w, zero),
		);
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
		frame.secondary = (v) =>
			plot.y + plot.h - chartAxisRatio(v, second, model.axes?.secondary) * plot.h;
		out.push(
			axisLine(
				model,
				model.axes?.secondary,
				'valueAxis',
				plot.x + plot.w,
				plot.y,
				plot.x + plot.w,
				plot.y + plot.h,
			),
		);
		if (
			model.axes?.secondary?.source.majorGridlines ||
			model.axes?.secondary?.source.minorGridlines
		)
			out.push(
				axisGridlines(model, model.axes.secondary, second.ticks, frame.secondary, plot, false),
			);
		if (chartAxisLabels(model.axes?.secondary))
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
		series: ChartViewModel['series'][number];
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
				radius: area ? 0 : chartMarkerRadius(s, i, 2.5),
			});
			lower.push([frame.band(i), frame.value(below)]);
		});
		layers.push({
			series: s,
			points,
			lower,
			color: s.color,
			filter: s.shadowFilter ? ` filter="${s.shadowFilter}"` : '',
		});
	});
	for (const { points, lower, color, filter, series } of area ? [...layers].reverse() : layers) {
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
		points.forEach(([x, y], index) => {
			if (!Number.isNaN(x)) out.push(chartSeriesMarkerSvg(series, index, x, y, 2.5));
		});
	}
	pointLabels(model, marks, out);
}
