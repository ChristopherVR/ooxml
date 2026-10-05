/**
 * Headless chart operations for chart-group spacing and geometry
 * (`c:gapWidth`, `c:overlap`, `c:firstSliceAng`, `c:holeSize`) and for the
 * chart-area / plot-area format (`c:spPr` fill and border,
 * `c:roundedCorners`). Work on new charts and on charts loaded from a file:
 * the save path rewrites only what differs from the authored XML. Split out
 * of `chart-operations.ts`, already well over the repo's 300-line guideline.
 *
 * @module sdk/chart-layout-operations
 */

import type { ChartPptxElement } from '../../types/elements';
import type { ChartAreaFormatInput, ChartGroupOptionsInput } from './chart-layout-input';
import { applyChartAreaFormatInput, applyChartGroupOptionsInput } from './chart-layout-input';
import { ensureChartData, setChartAxis } from './chart-operations';
import type { ChartInput } from './types';

/**
 * Set (or clear with `null`) bar spacing, slice angle and hole size.
 *
 * @throws {RangeError} A value outside its schema range (`ST_GapAmount` 0..500,
 *   `ST_Overlap` -100..100, `ST_FirstSliceAng` 0..360, `ST_HoleSize` 1..90) or
 *   not an integer.
 * @throws {Error} An option the chart type has no element for.
 *
 * @example
 * ```ts
 * setChartGroupOptions(barChart, { gapWidth: 35, overlap: 10 });
 * setChartGroupOptions(doughnut, { firstSliceAngle: 90, holeSize: 60 });
 * setChartGroupOptions(barChart, { overlap: null }); // back to the default
 * ```
 */
export function setChartGroupOptions(
	element: ChartPptxElement,
	options: ChartGroupOptionsInput,
): void {
	ensureChartData(element);
	applyChartGroupOptionsInput(element.chartData, options);
}

/**
 * Set the fill, border and (chart area only) rounded corners of the chart area
 * (`c:chartSpace/c:spPr`) or the plot area (`c:plotArea/c:spPr`). Pass `null`
 * to clear the area's format entirely.
 *
 * @throws {Error} A malformed colour, or `roundedCorners` on the plot area.
 * @throws {RangeError} An invalid gradient.
 *
 * @example
 * ```ts
 * // Transparent chart with no border and square corners
 * setChartAreaFormat(chartEl, "chart", { fill: "none", border: "none", roundedCorners: false });
 * setChartAreaFormat(chartEl, "plot", { fill: "#F2F2F2" });
 * ```
 */
export function setChartAreaFormat(
	element: ChartPptxElement,
	area: 'chart' | 'plot',
	format: ChartAreaFormatInput | null,
): void {
	ensureChartData(element);
	applyChartAreaFormatInput(element.chartData, area, format);
}

/**
 * Apply the layout fields of a {@link ChartInput} (group options, area
 * formats, axis edits) to a chart element `createChartElement` just built.
 */
export function applyChartInputLayout(element: ChartPptxElement, input: ChartInput): void {
	setChartGroupOptions(element, {
		...(input.gapWidth !== undefined ? { gapWidth: input.gapWidth } : {}),
		...(input.overlap !== undefined ? { overlap: input.overlap } : {}),
		...(input.firstSliceAngle !== undefined ? { firstSliceAngle: input.firstSliceAngle } : {}),
		...(input.holeSize !== undefined ? { holeSize: input.holeSize } : {}),
	});
	if (input.chartArea) {
		setChartAreaFormat(element, 'chart', input.chartArea);
	}
	if (input.plotArea) {
		setChartAreaFormat(element, 'plot', input.plotArea);
	}
	for (const [axisType, edit] of Object.entries(input.axes ?? {})) {
		if (edit) {
			setChartAxis(element, axisType as keyof NonNullable<ChartInput['axes']>, edit);
		}
	}
}
