// Small helpers of the conditional-format evaluator.
import type { CellValue, DifferentialStyle } from '../model.js';
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
