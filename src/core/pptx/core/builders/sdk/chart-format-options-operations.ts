/**
 * Headless operations behind the element-specific Options tabs of
 * PowerPoint's chart Format panes: Series Options (gap width, overlap, first
 * slice angle, hole size), Axis Options (bounds, units, crossing, number
 * format, tick marks, labels, display units, order, log scale), Legend
 * Options (position, overlap) and Label Options (contents, position). Each
 * is pure and returns a new `PptxChartData`; a `null` clears a value back to
 * PowerPoint's automatic setting.
 *
 * @module sdk/chart-format-options-operations
 */
import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartDataLabelOptions,
	PptxChartDataLabelPosition,
	PptxChartTickMark,
} from '../../types';
import type { ChartGroupOptionsInput } from './chart-layout-input';
import { applyChartGroupOptionsInput } from './chart-layout-input';

/** Series Options: chart-group spacing. Throws like `setChartGroupOptions` on a bad value. */
export function withChartSeriesOptions(
	data: PptxChartData,
	options: ChartGroupOptionsInput,
): PptxChartData {
	const next: PptxChartData = {
		...data,
		...(data.ofPieOptions ? { ofPieOptions: { ...data.ofPieOptions } } : {}),
	};
	applyChartGroupOptionsInput(next, options);
	return next;
}

/** Axis Options edit. `undefined` leaves a field alone; `null` clears it. */
export interface ChartAxisOptionsEdit {
	min?: number | null;
	max?: number | null;
	majorUnit?: number | null;
	minorUnit?: number | null;
	/** Where the other axis crosses: automatic, the minimum, the maximum, or a value. */
	crosses?: 'autoZero' | 'min' | 'max' | number;
	numberFormat?: string | null;
	majorTickMark?: PptxChartTickMark;
	minorTickMark?: PptxChartTickMark;
	tickLabelPosition?: 'high' | 'low' | 'nextTo' | 'none';
	reverseOrder?: boolean;
	logBase?: number | null;
	displayUnits?: PptxChartAxisFormatting['displayUnits'] | null;
	visible?: boolean;
}

function assign<K extends keyof PptxChartAxisFormatting>(
	axis: PptxChartAxisFormatting,
	key: K,
	value: PptxChartAxisFormatting[K] | null | undefined,
): void {
	if (value === undefined) {
		return;
	}
	if (value === null) {
		delete axis[key];
	} else {
		axis[key] = value;
	}
}

/** Apply an Axis Options edit to `data.axes[axisIndex]`. */
export function withChartAxisOptions(
	data: PptxChartData,
	axisIndex: number,
	edit: ChartAxisOptionsEdit,
): PptxChartData {
	const current = data.axes?.[axisIndex];
	if (!current) {
		return data;
	}
	const axis: PptxChartAxisFormatting = { ...current };
	for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const) {
		const value = edit[key];
		if (value !== undefined && value !== null && !Number.isFinite(value)) {
			throw new RangeError(`Axis ${key} must be a finite number.`);
		}
		assign(axis, key, value);
	}
	if (edit.crosses !== undefined) {
		if (typeof edit.crosses === 'number') {
			if (!Number.isFinite(edit.crosses)) {
				throw new RangeError('Axis crossing value must be a finite number.');
			}
			delete axis.crosses;
			axis.crossesAt = edit.crosses;
		} else {
			delete axis.crossesAt;
			axis.crosses = edit.crosses;
		}
	}
	if (edit.numberFormat !== undefined) {
		if (edit.numberFormat === null || edit.numberFormat === '') {
			axis.numFmt = { formatCode: 'General', sourceLinked: true };
		} else {
			axis.numFmt = { formatCode: edit.numberFormat, sourceLinked: false };
		}
	}
	assign(axis, 'majorTickMark', edit.majorTickMark);
	assign(axis, 'minorTickMark', edit.minorTickMark);
	assign(axis, 'tickLblPos', edit.tickLabelPosition);
	if (edit.reverseOrder !== undefined) {
		axis.orientation = edit.reverseOrder ? 'maxMin' : 'minMax';
	}
	if (edit.logBase !== undefined) {
		if (edit.logBase === null) {
			axis.logScale = false;
			delete axis.logBase;
		} else {
			if (!(edit.logBase >= 2 && edit.logBase <= 1000)) {
				throw new RangeError('Logarithmic base must be between 2 and 1000.');
			}
			axis.logScale = true;
			axis.logBase = edit.logBase;
		}
	}
	assign(axis, 'displayUnits', edit.displayUnits);
	if (edit.visible !== undefined) {
		axis.deleted = !edit.visible;
	}
	const axes = (data.axes ?? []).map((a, i) => (i === axisIndex ? axis : a));
	return { ...data, axes };
}

/** Legend Options edit. */
export interface ChartLegendOptionsEdit {
	show?: boolean;
	position?: 't' | 'b' | 'l' | 'r' | 'tr';
	/** "Show the legend without overlapping the chart" is `overlay: false`. */
	overlay?: boolean;
}

/** Apply a Legend Options edit. */
export function withChartLegendOptions(
	data: PptxChartData,
	edit: ChartLegendOptionsEdit,
): PptxChartData {
	const style = { ...data.style };
	if (edit.show !== undefined) {
		style.hasLegend = edit.show;
	}
	if (edit.position !== undefined) {
		style.legendPosition = edit.position;
		style.hasLegend = style.hasLegend ?? true;
	}
	if (edit.overlay !== undefined) {
		style.legendOverlay = edit.overlay;
	}
	return { ...data, style };
}

/** Label Options edit (chart-wide data labels). */
export interface ChartDataLabelOptionsEdit {
	show?: boolean;
	showValue?: boolean;
	showCategory?: boolean;
	showSeriesName?: boolean;
	showPercent?: boolean;
	showLegendKey?: boolean;
	position?: PptxChartDataLabelPosition | null;
}

/** Apply a Label Options edit to the chart-wide data labels. */
export function withChartDataLabelOptions(
	data: PptxChartData,
	edit: ChartDataLabelOptionsEdit,
): PptxChartData {
	const style = { ...data.style };
	const labels: PptxChartDataLabelOptions = { ...style.dataLabels };
	if (edit.show !== undefined) {
		style.hasDataLabels = edit.show;
	}
	for (const key of [
		'showValue',
		'showCategory',
		'showSeriesName',
		'showPercent',
		'showLegendKey',
	] as const) {
		if (edit[key] !== undefined) {
			labels[key] = edit[key];
			style.hasDataLabels = style.hasDataLabels ?? true;
		}
	}
	if (edit.position !== undefined) {
		if (edit.position === null) {
			delete labels.position;
		} else {
			labels.position = edit.position;
		}
	}
	style.dataLabels = labels;
	return { ...data, style };
}
