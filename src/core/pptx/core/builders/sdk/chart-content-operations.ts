/**
 * Applies the content fields of a {@link ChartInput} (`chart-content-input.ts`)
 * to a chart element `createChartElement` just built, through the operations
 * an edit uses: `setChartSeriesChartType` for combination series,
 * `withChartSeriesOnSecondaryAxis` for the secondary axis,
 * `setChartSeriesMarker`, `setChartDataPointFill` (a point marker's fill on
 * series drawn with markers), `withChartElementFormat`
 * for series fills and label text, and the data-label and font options
 * (`chart-content-labels.ts`).
 *
 * @module sdk/chart-content-operations
 */

import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartSeries,
	PptxChartType,
} from '../../types/chart';
import type { ChartPptxElement } from '../../types/elements';
import { canGenerateChartEx } from '../../utils/chart-cx-generator';
import { GENERATED_CHART_AXIS_IDS } from '../../utils/chart-xml-container-map';
import type { ChartSeriesContentInput, ChartSeriesMarkerInput } from './chart-content-input';
import { applyChartDataLabelsInput, applyChartFontsInput } from './chart-content-labels';
import { withChartElementFormat } from './chart-element-format-operations';
import { chartColorInput, chartLineWidthInput } from './chart-input-values';
import {
	ensureDataPoint,
	setChartDataPointFill,
	setChartSeriesChartType,
	setChartSeriesColor,
	setChartSeriesMarker,
} from './chart-operations';
import {
	chartSupportsSecondaryAxis,
	withChartSeriesOnSecondaryAxis,
} from './chart-series-axis-operations';
import type { ChartInput, ChartSeriesInput } from './types';

/** Chart types whose points vary in colour, so a palette colours each point. */
const VARIED = new Set<PptxChartType>(['pie', 'pie3D', 'doughnut', 'ofPie']);
const COMBO_BASE = new Set<PptxChartType>(['bar', 'line', 'area']);
const SMOOTHED = new Set<PptxChartType>(['line', 'line3D', 'scatter']);
const MARKED = new Set<PptxChartType>(['line', 'line3D', 'scatter', 'radar', 'stock']);
const XY = new Set<PptxChartType>(['scatter', 'bubble']);

/** The type a series is drawn as. */
export const drawnSeriesType = (data: PptxChartData, series: PptxChartSeries): PptxChartType =>
	series.seriesChartType ?? data.chartType;

/** Values with blanks: `null` is kept as a `0` placeholder flagged in `blanks`. */
export function chartSeriesValuesFromInput(
	values: readonly (number | null)[],
): Pick<PptxChartSeries, 'values' | 'blanks'> {
	const blanks = values.map((value) => value === null);
	return {
		values: values.map((value) => value ?? 0),
		...(blanks.includes(true) ? { blanks } : {}),
	};
}

function numbers(
	values: readonly number[] | undefined,
	count: number,
	name: string,
): number[] | undefined {
	if (values === undefined) return undefined;
	if (values.length !== count || values.some((value) => !Number.isFinite(value))) {
		throw new RangeError(`${name} must be ${count} finite numbers, one per value.`);
	}
	return [...values];
}

function requireType(allowed: ReadonlySet<PptxChartType>, type: PptxChartType, name: string) {
	if (!allowed.has(type)) {
		throw new Error(`${name} applies to ${[...allowed].join(', ')} series, not "${type}".`);
	}
}

/** Combination series types, set on every series once one differs from the chart's type. */
function applySeriesTypes(element: ChartPptxElement, inputs: readonly ChartSeriesInput[]): void {
	const base = element.chartData!.chartType;
	if (!inputs.some((input) => input.chartType !== undefined && input.chartType !== base)) return;
	if (!COMBO_BASE.has(base)) {
		throw new Error(`Series chart types combine bar, line and area charts, not "${base}".`);
	}
	inputs.forEach((input, index) => {
		const type = input.chartType ?? base;
		requireType(COMBO_BASE, type, 'A series chartType');
		setChartSeriesChartType(element, index, type);
	});
}

function marker(input: ChartSeriesMarkerInput) {
	const size = input.size;
	if (size !== undefined && (!Number.isInteger(size) || size < 2 || size > 72)) {
		throw new RangeError(`Marker size must be an integer from 2 to 72 points, got ${size}.`);
	}
	const line = input.line;
	const spPr = {
		...(input.fill !== undefined ? { fillColor: chartColorInput(input.fill, 'Marker fill') } : {}),
		...(line === 'none' ? { lineNoFill: true } : {}),
		...(line && line !== 'none' && line.color !== undefined
			? { strokeColor: chartColorInput(line.color, 'Marker line color') }
			: {}),
		...(line && line !== 'none' && line.width !== undefined
			? { strokeWidth: chartLineWidthInput(line.width, 'Marker line width') }
			: {}),
	};
	return { symbol: input.symbol, ...(size !== undefined ? { size } : {}), spPr };
}

/** Line, smoothing, marker, X values and bubble sizes of one series. */
function applySeriesShape(
	element: ChartPptxElement,
	index: number,
	input: ChartSeriesContentInput,
): void {
	const data = element.chartData!;
	const series = data.series[index]!;
	const type = drawnSeriesType(data, series);
	if (input.line === 'none') {
		series.lineNoFill = true;
	} else if (input.line) {
		if (input.line.width !== undefined) {
			series.lineWidth = chartLineWidthInput(input.line.width, 'Series line width');
		}
		if (input.line.dash !== undefined) series.lineDashStyle = input.line.dash;
	}
	if (input.smooth !== undefined) {
		requireType(SMOOTHED, type, 'smooth');
		series.smooth = input.smooth;
	}
	if (input.marker) {
		requireType(MARKED, type, 'A marker');
		setChartSeriesMarker(element, index, marker(input.marker));
	}
	const xValues = numbers(input.xValues, series.values.length, 'xValues');
	if (xValues) {
		requireType(XY, type, 'xValues');
		series.xValues = xValues;
	}
	const sizes = numbers(input.bubbleSizes, series.values.length, 'bubbleSizes');
	if (sizes) {
		requireType(new Set<PptxChartType>(['bubble']), type, 'bubbleSizes');
		series.bubbleSizes = sizes;
	}
}

/**
 * Whether a series' points are drawn as markers (line, scatter, stock, unfilled radar), so a
 * point colour is its marker's fill: PowerPoint draws no fill of a line point's own `c:spPr`.
 */
function paintsMarkers(data: PptxChartData, series: PptxChartSeries): boolean {
	const type = drawnSeriesType(data, series);
	if (type === 'radar') return data.radarStyle !== 'filled';
	return type === 'line' || type === 'scatter' || type === 'stock';
}

/** A point marker (`c:dPt/c:marker`) like the series' own, with its own fill. */
function setPointMarkerFill(series: PptxChartSeries, point: number, color: string): void {
	const own = series.marker;
	if (own?.symbol === 'none') {
		throw new Error('pointColors colour the markers of a line series, and this one has none.');
	}
	ensureDataPoint(series, point).marker = {
		symbol: own?.symbol ?? 'auto',
		...(own?.size !== undefined ? { size: own.size } : {}),
		spPr: { ...own?.spPr, fillColor: color },
	};
}

/** Series no-fill, per-point colours, and the palette. */
function applySeriesPaint(element: ChartPptxElement, input: ChartInput): void {
	const palette = input.palette?.map((color) => chartColorInput(color, 'Palette colour'));
	input.series.forEach((seriesInput, index) => {
		const data = element.chartData!;
		const series = data.series[index]!;
		if (seriesInput.fill === 'none') {
			if (seriesInput.color !== undefined || seriesInput.gradientFill !== undefined) {
				throw new Error("A series fill of 'none' cannot be combined with color or gradientFill.");
			}
			element.chartData = withChartElementFormat(
				data,
				{ kind: 'series', seriesIndex: index },
				{ fill: { kind: 'none' } },
			);
		}
		const pointColors = seriesInput.pointColors;
		if (pointColors && pointColors.length > series.values.length) {
			throw new RangeError(
				`pointColors has ${pointColors.length} entries for ${series.values.length} values.`,
			);
		}
		const own = pointColors?.map((color) =>
			color === null
				? undefined
				: color.trim() === 'none'
					? 'none'
					: chartColorInput(color, 'Point colour'),
		);
		const varied = VARIED.has(drawnSeriesType(data, series));
		// The series colour first: setChartSeriesColor clears the points' own paint.
		const owned = seriesInput.color ?? seriesInput.gradientFill ?? seriesInput.fill;
		if (!varied && palette?.length && owned === undefined) {
			setChartSeriesColor(element, index, palette[index % palette.length]);
		}
		const points = series.values.map(
			(_, point) =>
				own?.[point] ?? (varied && palette?.length ? palette[point % palette.length] : undefined),
		);
		const marked = paintsMarkers(data, series);
		points.forEach((color, point) => {
			if (!color) return;
			if (marked) setPointMarkerFill(series, point, color);
			else setChartDataPointFill(element, index, point, color);
		});
	});
}

/**
 * The generated axis ids on the primary axis pair (`GENERATED_CHART_AXIS_IDS`),
 * which `withChartSeriesOnSecondaryAxis` pairs a secondary axis with. Axes are
 * only appended, so `elementFormats` axis keys keep their indices.
 */
function withGeneratedAxisIds(data: PptxChartData): PptxChartData {
	const { cat, val } = GENERATED_CHART_AXIS_IDS;
	const axes: PptxChartAxisFormatting[] = (data.axes ?? []).map((axis) => ({ ...axis }));
	const xy = XY.has(data.chartType);
	const take = (type: 'catAx' | 'valAx', skip: number) => {
		const found = axes.filter((axis) => axis.axisType === type)[skip];
		if (found) return found;
		const axis: PptxChartAxisFormatting = { axisType: type };
		axes.push(axis);
		return axis;
	};
	const first = take(xy ? 'valAx' : 'catAx', 0);
	const second = take('valAx', xy ? 1 : 0);
	Object.assign(first, { axisId: first.axisId ?? cat, crossAxisId: first.crossAxisId ?? val });
	first.axPos ??= 'b';
	Object.assign(second, { axisId: second.axisId ?? val, crossAxisId: second.crossAxisId ?? cat });
	second.axPos ??= 'l';
	return { ...data, axes };
}

function applySecondaryAxes(element: ChartPptxElement, inputs: readonly ChartSeriesInput[]): void {
	if (!inputs.some((input) => input.secondaryAxis)) return;
	let data = withGeneratedAxisIds(element.chartData!);
	if (!chartSupportsSecondaryAxis(data)) {
		throw new Error(
			`A secondary axis applies to bar, line, area, scatter and bubble charts, not "${data.chartType}".`,
		);
	}
	inputs.forEach((input, index) => {
		if (input.secondaryAxis) data = withChartSeriesOnSecondaryAxis(data, index, true);
	});
	element.chartData = data;
}

/**
 * Apply the content fields of `input` to `element`.
 *
 * @throws {RangeError} A value outside its range, or a data label position the
 *   series' chart type does not allow.
 * @throws {Error} A field the chart type has no element for, or a malformed colour.
 */
export function applyChartContentInput(element: ChartPptxElement, input: ChartInput): void {
	const data = element.chartData!;
	if (input.dispBlanksAs !== undefined) {
		data.chartChrome = { ...data.chartChrome, dispBlanksAs: input.dispBlanksAs };
	}
	if (input.scatterStyle !== undefined) {
		requireType(new Set<PptxChartType>(['scatter']), data.chartType, 'scatterStyle');
		data.scatterStyle = input.scatterStyle;
	}
	if (input.embedWorkbook) {
		if (canGenerateChartEx(data)) {
			throw new Error(
				`embedWorkbook applies to classic charts; "${data.chartType}" is an extended chart.`,
			);
		}
		data.embedWorkbook = true;
	}
	applySeriesTypes(element, input.series);
	input.series.forEach((series, index) => applySeriesShape(element, index, series));
	applySeriesPaint(element, input);
	applySecondaryAxes(element, input.series);
	applyChartDataLabelsInput(element, input);
	applyChartFontsInput(element, input);
}
