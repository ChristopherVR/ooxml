/**
 * SDK input for chart-group spacing (`c:gapWidth`, `c:overlap`,
 * `c:firstSliceAng`, `c:holeSize`) and for the chart-area / plot-area format
 * (`c:spPr` fill and border, `c:roundedCorners`), with schema validation.
 * Shared by `createChartElement`, `ChartBuilder` and the `setChart*`
 * operations so every entry point applies the same rules.
 *
 * @module sdk/chart-layout-input
 */

import type { PptxChartData, PptxChartType } from '../../types/chart';
import type { ChartGroupOptionSpec } from '../../utils/chart-group-options';
import { CHART_GROUP_OPTIONS } from '../../utils/chart-group-options';
import {
	chartContainerAllows,
	chartTypeToContainerLocalName,
} from '../../utils/chart-container-content-model';
import type { ChartGradientInput } from './chart-gradient-input';
import { toChartGradientFill } from './chart-gradient-input';

/**
 * Chart-group spacing and geometry. `null` clears a value (a loaded chart
 * then drops the element; a new chart falls back to its default).
 */
export interface ChartGroupOptionsInput {
	/** Gap between bar clusters, % of bar width (`ST_GapAmount`, 0..500). Bar, bar3D, ofPie. */
	gapWidth?: number | null;
	/** Overlap of bars within a cluster, % (`ST_Overlap`, -100..100). Bar (2-D) only. */
	overlap?: number | null;
	/** First slice angle, degrees clockwise from 12 o'clock (`ST_FirstSliceAng`, 0..360). Pie, doughnut. */
	firstSliceAngle?: number | null;
	/** Doughnut hole size, % of the diameter (`ST_HoleSize`, 1..90). Doughnut only. */
	holeSize?: number | null;
}

/**
 * Chart-area or plot-area format. `fill` is a hex colour, `'none'` or a
 * gradient; `border` is a hex colour or `'none'`; `null` clears either.
 * `roundedCorners` applies to the chart area only (PowerPoint draws rounded
 * corners when it is unset, so pass `false` for square ones).
 */
export interface ChartAreaFormatInput {
	fill?: string | ChartGradientInput | null;
	border?: string | null;
	roundedCorners?: boolean | null;
}

const INPUT_KEYS: Record<ChartGroupOptionSpec['element'], keyof ChartGroupOptionsInput> = {
	gapWidth: 'gapWidth',
	overlap: 'overlap',
	firstSliceAng: 'firstSliceAngle',
	holeSize: 'holeSize',
};

const CLASSIC_TYPES: readonly PptxChartType[] = [
	'bar',
	'bar3D',
	'line',
	'line3D',
	'pie',
	'pie3D',
	'ofPie',
	'doughnut',
	'area',
	'area3D',
	'scatter',
	'bubble',
	'radar',
	'stock',
	'surface',
];

const HEX_COLOR = /^#?[0-9a-f]{6}$/iu;

function allowedTypes(element: string): PptxChartType[] {
	return CLASSIC_TYPES.filter((type) => {
		const local = chartTypeToContainerLocalName(type);
		return local !== undefined && chartContainerAllows(local, element);
	});
}

/** Chart types the chart is drawn with: its own, or each series' in a combo. */
function drawnTypes(chartData: PptxChartData): PptxChartType[] {
	if (chartData.chartType !== 'combo') {
		return [chartData.chartType];
	}
	return chartData.series.map((s) => s.seriesChartType ?? chartData.chartType);
}

function validateValue(spec: ChartGroupOptionSpec, name: string, value: number): void {
	if (!Number.isInteger(value) || value < spec.min || value > spec.max) {
		throw new RangeError(
			`${name} must be an integer from ${spec.min} to ${spec.max} (${spec.schemaType}), got ${value}.`,
		);
	}
}

/**
 * Validate `input` against the schema and the chart type, then apply it to
 * the model.
 *
 * @throws {RangeError} A value outside its schema range or not an integer.
 * @throws {Error} An option the chart type has no element for (for example
 *   `holeSize` on a bar chart).
 */
export function applyChartGroupOptionsInput(
	chartData: PptxChartData,
	input: ChartGroupOptionsInput,
): void {
	for (const spec of CHART_GROUP_OPTIONS) {
		const name = INPUT_KEYS[spec.element];
		const value = input[name];
		if (value === undefined) {
			continue;
		}
		if (value === null) {
			delete chartData[spec.field];
			continue;
		}
		validateValue(spec, name, value);
		const types = drawnTypes(chartData);
		const allowed = allowedTypes(spec.element);
		if (!types.some((type) => allowed.includes(type))) {
			throw new Error(
				`${name} applies to ${allowed.join(', ')} charts; this chart is drawn as "${types.join('", "')}".`,
			);
		}
		chartData[spec.field] = value;
		if (spec.element === 'gapWidth' && chartData.ofPieOptions) {
			chartData.ofPieOptions.gapWidth = value;
		}
	}
}

function toPaint(value: string, name: string): string {
	const trimmed = value.trim();
	if (trimmed === 'none') {
		return 'none';
	}
	if (!HEX_COLOR.test(trimmed)) {
		throw new Error(`${name} must be a 6-digit hex colour or 'none', got "${value}".`);
	}
	return `#${trimmed.replace(/^#/u, '').toUpperCase()}`;
}

/**
 * Validate `input` and apply it to the chart-area (`'chart'`) or plot-area
 * (`'plot'`) fields of the model's style; `null` clears every field.
 *
 * @throws {Error} A malformed colour, or `roundedCorners` on the plot area.
 * @throws {RangeError} An invalid gradient (see `toChartGradientFill`).
 */
export function applyChartAreaFormatInput(
	chartData: PptxChartData,
	area: 'chart' | 'plot',
	input: ChartAreaFormatInput | null,
): void {
	const style = (chartData.style ??= {});
	const fillKey = area === 'chart' ? 'chartAreaFill' : 'plotAreaFill';
	const gradientKey = area === 'chart' ? 'chartAreaGradient' : 'plotAreaGradient';
	const borderKey = area === 'chart' ? 'chartAreaBorder' : 'plotAreaBorder';
	if (input === null) {
		delete style[fillKey];
		delete style[gradientKey];
		delete style[borderKey];
		if (area === 'chart') {
			delete chartData.roundedCorners;
		}
		return;
	}
	if (input.roundedCorners !== undefined && area !== 'chart') {
		throw new Error('roundedCorners applies to the chart area only.');
	}
	if (input.fill !== undefined) {
		delete style[fillKey];
		delete style[gradientKey];
		if (typeof input.fill === 'string') {
			style[fillKey] = toPaint(input.fill, 'fill');
		} else if (input.fill !== null) {
			style[gradientKey] = toChartGradientFill(input.fill);
		}
	}
	if (input.border !== undefined) {
		if (input.border === null) {
			delete style[borderKey];
		} else {
			style[borderKey] = toPaint(input.border, 'border');
		}
	}
	if (input.roundedCorners !== undefined) {
		if (input.roundedCorners === null) {
			delete chartData.roundedCorners;
		} else {
			chartData.roundedCorners = input.roundedCorners;
		}
	}
}
