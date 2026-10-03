// Small helpers of the conditional-format evaluator.
import type { CellValue, CfvoThreshold, DifferentialStyle } from '../model.js';
import { mixColors } from './colors.js';
import { dateToSerial } from '../numfmt/date.js';
import { mergeBorder, mergeFont } from './style-view.js';

/** A literal number or string operand, evaluated without the formula engine. */
export function literal(formula: string): CellValue | undefined {
	const text = formula.trim();
	if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(text)) return Number(text);
	const quoted = /^"((?:[^"]|"")*)"$/.exec(text);
	if (quoted) return (quoted[1] ?? '').replace(/""/g, '"');
	if (/^TRUE$/i.test(text)) return true;
	if (/^FALSE$/i.test(text)) return false;
	return undefined;
}

/** The serial of today's local date. */
export function localTodaySerial(date1904: boolean): number {
	const now = new Date();
	return Math.floor(
		dateToSerial(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())), date1904),
	);
}

/** Later (lower-priority) differential style under the accumulated one. */
export function mergeDxf(
	acc: DifferentialStyle | undefined,
	next: DifferentialStyle,
): DifferentialStyle {
	if (!acc) return next;
	const out: DifferentialStyle = { ...acc };
	if (next.font) out.font = mergeFont(next.font, acc.font);
	if (next.border) out.border = mergeBorder(next.border, acc.border);
	if (!acc.fill && next.fill) out.fill = next.fill;
	if (acc.numFmt === undefined && next.numFmt !== undefined) out.numFmt = next.numFmt;
	return out;
}

/**
 * The icon a value earns: the last threshold it reaches, where a threshold with `gte: false`
 * must be exceeded (`>`) rather than met (`>=`).
 */
export function iconIndex(
	value: number,
	cuts: readonly number[],
	list: readonly CfvoThreshold[],
): number {
	let index = 0;
	for (let i = 1; i < cuts.length; i++) {
		const cut = cuts[i] ?? Infinity;
		if (list[i]?.gte === false ? value > cut : value >= cut) index = i;
	}
	return index;
}

/** The colour-scale colour of `value` between threshold positions. */
export function scaleColor(
	value: number,
	cuts: readonly number[],
	colors: readonly string[],
): string | undefined {
	if (!cuts.length || cuts.length !== colors.length) return undefined;
	if (value <= (cuts[0] ?? 0)) return colors[0];
	const last = cuts.length - 1;
	if (value >= (cuts[last] ?? 0)) return colors[last];
	for (let i = 0; i < last; i++) {
		const lo = cuts[i] ?? 0;
		const hi = cuts[i + 1] ?? 0;
		if (value >= lo && value <= hi) {
			const a = colors[i] ?? '#FFFFFF';
			const b = colors[i + 1] ?? a;
			return hi > lo ? mixColors(a, b, (value - lo) / (hi - lo)) : b;
		}
	}
	return colors[last];
}
