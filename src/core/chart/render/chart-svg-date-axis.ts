import { calendarTicks, excelSerialToDate } from '../date-axis';
import { formatDate } from '../date-format';
import type { ChartViewModel } from './chart-view';
import type { Rect } from './chart-svg-util';
import { text } from './chart-svg-util';
import { chartTextAttributes } from './chart-svg-appearance';
import { strokeAttributes } from './chart-svg-analysis';

/** A date axis keeps source indices while ordering its marks on a continuous calendar scale. */
export function chartDateAxis(model: ChartViewModel, plot: Rect, out: string[]) {
	const axis = model.axes?.category?.source;
	const values = model.dateCategories;
	if (axis?.kind !== 'date' || !values?.length) return undefined;
	const points = values
		.flatMap((value, index) =>
			value === null || !Number.isFinite(value) ? [] : [{ value, index }],
		)
		.sort((a, b) => a.value - b.value);
	if (!points.length) return undefined;
	const min = axis.scaling.min ?? points[0]!.value;
	const max = axis.scaling.max ?? points.at(-1)!.value;
	if (!Number.isFinite(min) || !Number.isFinite(max) || max < min) return undefined;
	const position = (value: number) => {
		const ratio = max === min ? 0.5 : (value - min) / (max - min);
		const t = axis.scaling.orientation === 'maxMin' ? 1 - ratio : ratio;
		return model.horizontal ? plot.y + plot.h * (1 - t) : plot.x + plot.w * t;
	};
	const visible = points.filter((p) => p.value >= min && p.value <= max);
	const unitOf = (unit: string | undefined): 'days' | 'months' | 'years' =>
		unit === 'months' || unit === 'years' ? unit : 'days';
	const unit = unitOf(axis.majorTimeUnit ?? axis.baseTimeUnit);
	const ticks = calendarTicks(
		min,
		max,
		model.date1904 ?? false,
		unit,
		Math.max(1, axis.majorUnit ?? 1),
	);
	const attrs = chartTextAttributes(model, 'categoryAxis')!;
	const size = attrs.size ?? 10;
	const y = axis.position === 't' ? plot.y : plot.y + plot.h;
	const paint = model.axes?.category?.line;
	const stroke = { color: paint?.color ?? '#BFBFBF', width: paint?.width ?? 1 };
	const mark = (serial: number, placement: string | undefined, length: number) => {
		if (!placement || placement === 'none' || axis.deleted) return;
		const inward = axis.position === 't' ? length : -length;
		const start = placement === 'cross' ? -length : 0;
		const end = placement === 'cross' ? length : placement === 'in' ? inward : -inward;
		if (model.horizontal) {
			const x = axis.position === 'r' ? plot.x + plot.w : plot.x;
			const direction = axis.position === 'r' ? -1 : 1;
			const a = placement === 'cross' ? -length : 0;
			const b =
				placement === 'cross'
					? length
					: placement === 'in'
						? direction * length
						: -direction * length;
			out.push(
				`<line x1="${x + a}" y1="${position(serial)}" x2="${x + b}" y2="${position(serial)}" ${strokeAttributes(stroke)} data-chart-part="dateTick"/>`,
			);
		} else
			out.push(
				`<line x1="${position(serial)}" y1="${y + start}" x2="${position(serial)}" y2="${y + end}" ${strokeAttributes(stroke)} data-chart-part="dateTick"/>`,
			);
	};
	ticks.forEach((serial, index) => {
		if (index % Math.max(1, axis.tickMarkSkip ?? 1) === 0) mark(serial, axis.majorTickMark, 4);
		if (
			axis.deleted ||
			model.dataTable ||
			axis.tickLabelPosition === 'none' ||
			index % Math.max(1, axis.tickLabelSkip ?? 1)
		)
			return;
		const label = formatDate(
			excelSerialToDate(serial, model.date1904),
			unit,
			axis.numberFormat?.formatCode,
		);
		if (model.horizontal)
			out.push(text(plot.x - 4, position(serial) + size * 0.3, label, { ...attrs, anchor: 'end' }));
		else {
			const above =
				axis.tickLabelPosition === 'high' ||
				(axis.tickLabelPosition !== 'low' && axis.position === 't');
			const baseline =
				axis.tickLabelPosition === 'high'
					? plot.y
					: axis.tickLabelPosition === 'low'
						? plot.y + plot.h
						: y;
			out.push(
				text(position(serial), baseline + (above ? -4 : size + 3), label, {
					...attrs,
					anchor: 'middle',
				}),
			);
		}
	});
	if (axis.minorTickMark && axis.minorTickMark !== 'none')
		for (const serial of calendarTicks(
			min,
			max,
			model.date1904 ?? false,
			unitOf(axis.minorTimeUnit ?? axis.baseTimeUnit ?? unit),
			Math.max(1, axis.minorUnit ?? 1),
		))
			mark(serial, axis.minorTickMark, 2.5);
	return {
		order: visible.map((p) => p.index),
		band: (index: number) => position(values[index] ?? min),
		bandSize: (model.horizontal ? plot.h : plot.w) / Math.max(1, max - min + 1),
	};
}
