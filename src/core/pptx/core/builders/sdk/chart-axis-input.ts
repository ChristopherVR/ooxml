/**
 * The appearance fields of a `ChartAxisEdit` (`setChartAxis`, `ChartInput.axes`)
 * that the chart model records as Format-pane edits of the axis: tick label
 * rotation, italic labels, the axis line and the gridline lines. Mapped through
 * `withChartElementFormat`, so a new chart and a loaded one save them the same
 * way (`c:txPr/a:bodyPr/@rot`, `a:defRPr/@i`, `c:spPr/a:ln`, `c:majorGridlines/c:spPr/a:ln`).
 *
 * @module sdk/chart-axis-input
 */

import type { PptxChartData, PptxChartElementLine } from '../../types';
import type { StrokeDashType } from '../../types/common';
import { withChartElementFormat } from './chart-element-format-operations';
import type { ChartElementFormatPatch } from './chart-element-format-operations';
import { chartColorInput, chartLineWidthInput, chartTextRotationInput } from './chart-input-values';

/**
 * An axis or gridline line: a colour, width in points and preset dash (`a:prstDash`), or
 * `'none'` to hide the line. A width or dash without a colour keeps the line's colour.
 */
export type ChartAxisLineInput =
	| { color?: string; width?: number; dash?: Exclude<StrokeDashType, 'custom'> }
	| 'none';

/** The axis appearance fields `withChartAxisAppearance` writes. */
export interface ChartAxisAppearanceEdit {
	/** Tick label angle in degrees, -90..90; `null` clears it. */
	rotation?: number | null;
	/** Italic tick labels. */
	italic?: boolean;
	line?: ChartAxisLineInput;
	/** Major gridline line (`c:majorGridlines/c:spPr/a:ln`). */
	majorGridlinesLine?: ChartAxisLineInput;
	/** Minor gridline line (`c:minorGridlines/c:spPr/a:ln`). */
	minorGridlinesLine?: ChartAxisLineInput;
}

function axisLine(line: ChartAxisLineInput): Partial<PptxChartElementLine> {
	if (line === 'none') {
		return { kind: 'none' };
	}
	// A width or dash without a colour keeps the line's paint (automatic on a new chart).
	const out: Partial<PptxChartElementLine> = line.color !== undefined ? { kind: 'solid' } : {};
	if (line.color !== undefined) out.color = chartColorInput(line.color, 'axis line color');
	if (line.width !== undefined) out.width = chartLineWidthInput(line.width, 'axis line width');
	if (line.dash !== undefined) {
		if ((line.dash as string) === 'custom')
			throw new Error('An axis line dash must be a preset dash.');
		out.dash = line.dash;
	}
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
	let next = data;
	if (patch.text || patch.line) {
		next = withChartElementFormat(next, { kind: 'axis', axisIndex }, patch);
	}
	for (const kind of ['majorGridlines', 'minorGridlines'] as const) {
		const gridline = edit[`${kind}Line`];
		if (gridline !== undefined)
			next = withChartElementFormat(next, { kind, axisIndex }, { line: axisLine(gridline) });
	}
	return next;
}
