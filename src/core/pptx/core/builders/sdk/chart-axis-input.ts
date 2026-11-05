/**
 * The appearance fields of a `ChartAxisEdit` (`setChartAxis`, `ChartInput.axes`)
 * that the chart model records as Format-pane edits of the axis: tick label
 * rotation, italic labels and the axis line. Mapped through
 * `withChartElementFormat`, so a new chart and a loaded one save them the same
 * way (`c:txPr/a:bodyPr/@rot`, `a:defRPr/@i`, `c:spPr/a:ln`).
 *
 * @module sdk/chart-axis-input
 */

import type { PptxChartData, PptxChartElementLine } from '../../types';
import { withChartElementFormat } from './chart-element-format-operations';
import type { ChartElementFormatPatch } from './chart-element-format-operations';
import { chartColorInput, chartLineWidthInput, chartTextRotationInput } from './chart-input-values';

/**
 * Axis line: a colour and width in points, or `'none'` to hide the line. A width without a
 * colour keeps the line's colour.
 */
export type ChartAxisLineInput = { color?: string; width?: number } | 'none';

/** The axis appearance fields `withChartAxisAppearance` writes. */
export interface ChartAxisAppearanceEdit {
	/** Tick label angle in degrees, -90..90; `null` clears it. */
	rotation?: number | null;
	/** Italic tick labels. */
	italic?: boolean;
	line?: ChartAxisLineInput;
}

function axisLine(line: ChartAxisLineInput): Partial<PptxChartElementLine> {
	if (line === 'none') {
		return { kind: 'none' };
	}
	// A width without a colour keeps the line's paint (automatic on a new chart).
	const out: Partial<PptxChartElementLine> = line.color !== undefined ? { kind: 'solid' } : {};
	if (line.color !== undefined) out.color = chartColorInput(line.color, 'axis line color');
	if (line.width !== undefined) out.width = chartLineWidthInput(line.width, 'axis line width');
	return out;
}

/**
 * Record the appearance fields of `edit` for `data.axes[axisIndex]`.
 *
 * @throws {RangeError} A rotation outside -90..90, a negative or non-finite width.
 * @throws {Error} A malformed colour.
 */
export function withChartAxisAppearance(
	data: PptxChartData,
	axisIndex: number,
	edit: ChartAxisAppearanceEdit,
): PptxChartData {
	const patch: ChartElementFormatPatch = {};
	if (edit.rotation !== undefined || edit.italic !== undefined) {
		patch.text = {
			...(edit.rotation !== undefined
				? { rotation: edit.rotation === null ? null : chartTextRotationInput(edit.rotation) }
				: {}),
			...(edit.italic !== undefined ? { italic: edit.italic } : {}),
		};
	}
	if (edit.line !== undefined) {
		patch.line = axisLine(edit.line);
	}
	if (!patch.text && !patch.line) {
		return data;
	}
	return withChartElementFormat(data, { kind: 'axis', axisIndex }, patch);
}
