/**
 * Validation of the scalar values of the chart SDK inputs (colours, widths,
 * sizes, angles and fonts), shared by `chart-content-operations.ts` and
 * `chart-axis-input.ts` so every entry point applies the same rules.
 *
 * @module sdk/chart-input-values
 */

import type { PptxChartLegendTextStyle } from '../../types/chart';
import type { ChartFontInput } from './chart-content-input';
import { toPaint } from './chart-layout-input';

/** A hex colour as `#RRGGBB`. @throws {Error} A malformed colour or `'none'`. */
export function chartColorInput(value: string, name: string): string {
	const paint = toPaint(value, name);
	if (paint === 'none') {
		throw new Error(`${name} must be a 6-digit hex colour, got "${value}".`);
	}
	return paint;
}

/** A line width in points. @throws {RangeError} A negative or non-finite width. */
export function chartLineWidthInput(value: number, name: string): number {
	if (!Number.isFinite(value) || value < 0 || value > 1584) {
		throw new RangeError(`${name} must be a width in points from 0 to 1584, got ${value}.`);
	}
	return value;
}

/** A text angle in degrees. @throws {RangeError} Outside -90..90 (PowerPoint's custom angle). */
export function chartTextRotationInput(value: number): number {
	if (!Number.isFinite(value) || value < -90 || value > 90) {
		throw new RangeError(`Text rotation must be from -90 to 90 degrees, got ${value}.`);
	}
	return value;
}

/** A font size in points. @throws {RangeError} Outside 1..400 (`ST_TextFontSize`). */
function fontSize(value: number): number {
	if (!Number.isFinite(value) || value < 1 || value > 400) {
		throw new RangeError(`Font size must be from 1 to 400 points, got ${value}.`);
	}
	return value;
}

/** A {@link ChartFontInput} as the model's chart text style. */
export function chartFontInput(font: ChartFontInput, name: string): PptxChartLegendTextStyle {
	const out: PptxChartLegendTextStyle = {};
	if (font.family !== undefined) out.fontFamily = font.family;
	if (font.size !== undefined) out.fontSize = fontSize(font.size);
	if (font.bold !== undefined) out.bold = font.bold;
	if (font.italic !== undefined) out.italic = font.italic;
	if (font.color !== undefined) out.color = chartColorInput(font.color, `${name} color`);
	return out;
}
