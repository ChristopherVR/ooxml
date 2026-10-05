/**
 * SDK input shape for chart gradient fills, and its validated conversion to
 * the {@link PptxChartGradientFill} model the save pipeline writes as
 * `c:spPr/a:gradFill` (see `utils/chart-gradient-fill-writer.ts`).
 *
 * @module sdk/chart-gradient-input
 */

import { OOXML_PERCENT_UNITS } from '../../constants';
import type { PptxChartGradientFill, PptxChartSeries, PptxChartType } from '../../types/chart';
import {
	DEFAULT_CHART_GRADIENT_ANGLE,
	supportsChartGradientFill,
} from '../../utils/chart-gradient-fill-writer';

/** One gradient stop. */
export interface ChartGradientStopInput {
	/** Hex colour, `#RRGGBB` or `RRGGBB`. */
	color: string;
	/** Stop position along the gradient, 0 to 100 (percent). */
	position: number;
	/** Stop opacity, 0 (transparent) to 1 (opaque, the default); written as `a:alpha`. */
	opacity?: number;
}

/**
 * A chart series or data-point gradient fill: at least two stops, linear
 * (the default) or radial.
 *
 * `angle` follows the OOXML `a:lin/@ang` convention, the same as the shape
 * `FillInput`: degrees clockwise from the positive x-axis, pointing from the
 * first stop towards the last (`0` = left to right, `90` = top to bottom, the
 * default). A radial gradient radiates from `focalPoint` (fractions 0 to 1 of
 * the mark's box, the centre by default) and is written as
 * `a:path path="circle"` with a matching `a:fillToRect`.
 *
 * @example
 * ```ts
 * const topToBottom: ChartGradientInput = {
 *   angle: 90,
 *   stops: [
 *     { color: "#60A5FA", position: 0 },
 *     { color: "#1E3A8A", position: 100 },
 *   ],
 * };
 * const glow: ChartGradientInput = {
 *   type: "radial",
 *   focalPoint: { x: 0.5, y: 0.5 },
 *   stops: [
 *     { color: "#FFFFFF", position: 0, opacity: 0.6 },
 *     { color: "#ED7D31", position: 100 },
 *   ],
 * };
 * ```
 */
export type ChartGradientInput =
	| { type?: 'linear'; angle?: number; stops: ChartGradientStopInput[] }
	| { type: 'radial'; focalPoint?: { x: number; y: number }; stops: ChartGradientStopInput[] };

const HEX_COLOR = /^#?[0-9a-f]{6}$/iu;
const FULL_TURN = 360;
/** Model positions are 0..100, so one `a:gs/@pos` unit is 1/1000 of a position. */
const POSITION_STEP = OOXML_PERCENT_UNITS / 100;
/** Opacity 0..1 maps onto `a:alpha/@val` 0..100000. */
const OPACITY_STEP = OOXML_PERCENT_UNITS;

function inRange(value: number, min: number, max: number): boolean {
	return Number.isFinite(value) && value >= min && value <= max;
}

function toStop(
	stop: ChartGradientStopInput,
	index: number,
): PptxChartGradientFill['stops'][number] {
	if (typeof stop.color !== 'string' || !HEX_COLOR.test(stop.color.trim())) {
		throw new Error(
			`Gradient stop ${index} colour must be a 6-digit hex string, got "${stop.color}".`,
		);
	}
	if (!inRange(stop.position, 0, 100)) {
		throw new RangeError(`Gradient stop ${index} position must be between 0 and 100.`);
	}
	if (stop.opacity !== undefined && !inRange(stop.opacity, 0, 1)) {
		throw new RangeError(`Gradient stop ${index} opacity must be between 0 and 1.`);
	}
	const color = `#${stop.color.trim().replace(/^#/u, '').toUpperCase()}`;
	const position = Math.round(stop.position * POSITION_STEP) / POSITION_STEP;
	const opacity =
		stop.opacity === undefined ? undefined : Math.round(stop.opacity * OPACITY_STEP) / OPACITY_STEP;
	return opacity !== undefined && opacity < 1 ? { color, position, opacity } : { color, position };
}

/**
 * Validate a {@link ChartGradientInput} and convert it to the chart model,
 * normalised the way a reload parses it back: stops sorted by position,
 * colours as `#RRGGBB`, a linear angle in `[0, 360)` (default 90), a radial
 * focal point defaulting to the centre.
 *
 * @throws {RangeError} Fewer than two stops, or a position, opacity, angle or
 *   focal point out of range.
 * @throws {Error} A stop colour that is not a 6-digit hex string.
 */
export function toChartGradientFill(input: ChartGradientInput): PptxChartGradientFill {
	if (!Array.isArray(input.stops) || input.stops.length < 2) {
		throw new RangeError('A chart gradient needs at least two stops.');
	}
	const stops = input.stops.map(toStop).sort((a, b) => a.position - b.position);
	if (input.type === 'radial') {
		const focalPoint = input.focalPoint ?? { x: 0.5, y: 0.5 };
		if (!inRange(focalPoint.x, 0, 1) || !inRange(focalPoint.y, 0, 1)) {
			throw new RangeError('A radial gradient focal point must lie within 0..1 on both axes.');
		}
		return { type: 'radial', stops, focalPoint: { x: focalPoint.x, y: focalPoint.y } };
	}
	const angle = input.angle ?? DEFAULT_CHART_GRADIENT_ANGLE;
	if (!Number.isFinite(angle)) {
		throw new RangeError('A linear gradient angle must be a finite number of degrees.');
	}
	return { type: 'linear', stops, angle: ((angle % FULL_TURN) + FULL_TURN) % FULL_TURN };
}

/**
 * Throw when a series drawn as `chartType` cannot carry a gradient fill:
 * line-drawn families (line, line3D, scatter, radar, stock) have no fillable
 * area, and ChartEx kinds (waterfall, funnel, treemap, ...) are not written
 * through `c:spPr`.
 *
 * @param where - Names the series in the error message, e.g. `"series 1"`.
 */
export function assertChartGradientSupported(chartType: PptxChartType, where: string): void {
	if (!supportsChartGradientFill(chartType)) {
		throw new Error(
			`Gradient fills apply to area-filled chart types (bar, area, pie, doughnut, ofPie, bubble, surface); ${where} is drawn as "${chartType}".`,
		);
	}
}

/**
 * The `gradientFill` to spread onto a new series built from a
 * `ChartSeriesInput.gradientFill` (`createChartElement`), validated against
 * the chart type it will be drawn as.
 */
export function chartSeriesGradientFromInput(
	chartType: PptxChartType,
	gradient: ChartGradientInput | undefined,
	seriesIndex: number,
): Pick<PptxChartSeries, 'gradientFill'> {
	if (!gradient) {
		return {};
	}
	assertChartGradientSupported(chartType, `series ${seriesIndex}`);
	return { gradientFill: toChartGradientFill(gradient) };
}
