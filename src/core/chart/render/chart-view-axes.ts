import type { ChartAxis, ChartShapeProperties } from '../model';
import type { DrawingColor } from '../../drawingml/types';
import { EMU_PER_PIXEL } from '../../units/constants';
import { computeLogValueRange, withAxisBounds } from '../axis-range';
import { generateAxisTicks, generateMinorAxisTicks } from '../axis-ticks';
import type { ChartAxisOptions, ChartAxisRange } from '../axis-range-types';
import type { AxisScale } from './chart-scale';
import type { ChartSummary } from './summary';
import type { ChartViewModel } from './chart-view';

export interface ChartAxisPaint {
	color?: string;
	width?: number;
	dash?: string;
}
export interface ChartAxisView {
	source: ChartAxis;
	range: ChartAxisRange;
	minorTicks: number[];
	line: ChartAxisPaint;
	majorGrid: ChartAxisPaint;
	minorGrid: ChartAxisPaint;
}
export interface ChartAxesView {
	category?: ChartAxisView;
	value?: ChartAxisView;
	secondary?: ChartAxisView;
}

/** A normalised position shared by the SVG painters, including reversed and log axes. */
export function chartAxisRatio(value: number, scale: AxisScale, axis?: ChartAxisView): number {
	const range = axis?.range;
	const ratio = range?.logScale
		? Math.log(Math.max(value, range.min) / range.min) / Math.log(range.max / range.min)
		: (value - scale.min) / (scale.max - scale.min || 1);
	return range?.reverseOrder ? 1 - ratio : ratio;
}

export function chartAxisLabels(axis?: ChartAxisView): boolean {
	return axis?.source.deleted !== true && axis?.source.tickLabelPosition !== 'none';
}

function options(axis: ChartAxis): ChartAxisOptions {
	return {
		...axis.scaling,
		...(axis.scaling.logBase === undefined ? {} : { logScale: true }),
		...(axis.majorUnit === undefined ? {} : { majorUnit: axis.majorUnit }),
		...(axis.minorUnit === undefined ? {} : { minorUnit: axis.minorUnit }),
	};
}

/** Bind parsed axes by plot-group IDs, never by their incidental XML order. */
export function applyChartAxes(
	model: ChartViewModel,
	chart: ChartSummary,
	color: (value: DrawingColor | undefined) => string | undefined,
): void {
	if (!chart.axes?.length) return;
	const axes = chart.axes;
	const linked = axes.filter((axis) => axis.id !== undefined && chart.axisIds?.includes(axis.id));
	const candidates = linked.length ? linked : axes;
	const xy = model.type === 'scatter' || model.type === 'bubble';
	const value =
		candidates.find((axis) => axis.kind === 'val' && !['b', 't'].includes(axis.position ?? '')) ??
		candidates.find((axis) => axis.kind === 'val');
	const category = candidates.find((axis) =>
		xy ? axis.kind === 'val' && axis !== value : axis.kind === 'cat' || axis.kind === 'date',
	);
	const secondary = axes.find((axis) => axis.kind === 'val' && axis !== value && axis !== category);
	const paint = (shape?: ChartShapeProperties): ChartAxisPaint => {
		const ln = shape?.line;
		const fill = ln?.fill;
		const resolved =
			fill?.kind === 'none' ? 'none' : color(fill?.kind === 'solid' ? fill.color : undefined);
		return {
			...(resolved === undefined ? {} : { color: resolved }),
			...(ln?.widthEmu === undefined ? {} : { width: ln.widthEmu / EMU_PER_PIXEL }),
			...(ln?.dash === undefined ? {} : { dash: ln.dash }),
		};
	};
	const bind = (axis: ChartAxis, scale?: AxisScale, values?: (number | null)[]): ChartAxisView => {
		const opts = options(axis);
		const automatic: ChartAxisRange = opts.logScale
			? computeLogValueRange(
					[{ values: values ?? [] }],
					opts.logBase && Number.isFinite(opts.logBase) && opts.logBase > 1 ? opts.logBase : 10,
				)
			: {
					min: scale?.min ?? 0,
					max: scale?.max ?? 1,
					span: (scale?.max ?? 1) - (scale?.min ?? 0),
					...(scale ? { majorUnit: scale.majorUnit } : {}),
				};
		const logBase = automatic.logBase;
		const range = withAxisBounds(automatic, opts, logBase);
		if (scale) {
			scale.min = range.min;
			scale.max = range.max;
			scale.ticks = generateAxisTicks(range, opts, 5);
		}
		return {
			source: axis,
			range,
			minorTicks: generateMinorAxisTicks(range, opts),
			line: paint(axis.spPr),
			majorGrid: paint(axis.majorGridlinesSpPr),
			minorGrid: paint(axis.minorGridlinesSpPr),
		};
	};
	const out: ChartAxesView = {};
	if (value)
		out.value = bind(
			value,
			model.valueAxis,
			[
				...model.series,
				...(model.extraGroups?.filter((g) => !g.secondary).flatMap((g) => g.series) ?? []),
			].flatMap((s) => s.values),
		);
	if (category)
		out.category = bind(
			category,
			model.xAxis,
			model.series.flatMap((s) => s.xValues ?? []),
		);
	if (secondary)
		out.secondary = bind(
			secondary,
			model.secondaryAxis,
			model.extraGroups
				?.filter((g) => g.secondary)
				.flatMap((g) => g.series.flatMap((s) => s.values)),
		);
	model.axes = out;
}
