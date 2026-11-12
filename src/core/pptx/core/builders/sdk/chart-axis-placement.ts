/**
 * The tick, crossing and gridline fields of a `ChartAxisEdit` (`setChartAxis`,
 * `ChartInput.axes`), checked and recorded on the axis entry: tick marks
 * (`c:majorTickMark`, `c:minorTickMark`), where the crossing axis meets it
 * (`c:crosses` or `c:crossesAt`), `c:crossBetween`, and whether a gridline line
 * shows or hides its gridlines.
 *
 * @module sdk/chart-axis-placement
 */

import type { PptxChartAxisFormatting } from '../../types/chart';
import type { PptxChartTickMark } from '../../types/chart-axis';
import type { ChartAxisLineInput } from './chart-axis-input';

const TICK_MARKS = new Set<string>(['cross', 'in', 'none', 'out']);

/** The placement fields of an axis edit. */
export interface ChartAxisPlacementEdit {
	majorTickMark?: PptxChartTickMark;
	minorTickMark?: PptxChartTickMark;
	crosses?: 'autoZero' | 'min' | 'max' | number;
	crossBetween?: 'between' | 'midCat';
	majorGridlines?: boolean;
	minorGridlines?: boolean;
	majorGridlinesLine?: ChartAxisLineInput;
	minorGridlinesLine?: ChartAxisLineInput;
}

/**
 * Record the tick marks and crossing of `edit` on `axis`.
 *
 * @throws {RangeError} A tick mark, crossing or `crossBetween` value outside its schema type.
 */
export function applyChartAxisPlacement(
	axis: PptxChartAxisFormatting,
	edit: ChartAxisPlacementEdit,
): void {
	for (const key of ['majorTickMark', 'minorTickMark'] as const) {
		const mark = edit[key];
		if (mark === undefined) continue;
		if (!TICK_MARKS.has(mark))
			throw new RangeError(`${key} must be cross, in, none or out, got "${mark}".`);
		axis[key] = mark;
	}
	if (edit.crosses !== undefined) {
		if (typeof edit.crosses === 'number') {
			if (!Number.isFinite(edit.crosses))
				throw new RangeError(`crosses must be a finite number, got ${edit.crosses}.`);
			axis.crossesAt = edit.crosses;
			delete axis.crosses;
		} else {
			if (!['autoZero', 'min', 'max'].includes(edit.crosses))
				throw new RangeError(
					`crosses must be autoZero, min, max or a number, got "${edit.crosses}".`,
				);
			axis.crosses = edit.crosses;
			delete axis.crossesAt;
		}
	}
	if (edit.crossBetween !== undefined) {
		if (edit.crossBetween !== 'between' && edit.crossBetween !== 'midCat')
			throw new RangeError(`crossBetween must be between or midCat, got "${edit.crossBetween}".`);
		axis.crossBetween = edit.crossBetween;
	}
}

/**
 * The gridline line of `kind` to format, if any: a line shows its gridlines, and `'none'` hides
 * them (and formats nothing) unless they are asked for as well.
 */
export function chartGridlinesLine(
	axis: PptxChartAxisFormatting,
	edit: ChartAxisPlacementEdit,
	kind: 'majorGridlines' | 'minorGridlines',
): ChartAxisLineInput | undefined {
	const line = edit[`${kind}Line`];
	if (line === undefined) return undefined;
	if (line !== 'none') axis[kind] ??= true;
	else if (edit[kind] !== true) {
		axis[kind] = false;
		return undefined;
	}
	return line;
}
