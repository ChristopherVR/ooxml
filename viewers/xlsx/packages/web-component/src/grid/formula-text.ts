// Pure helpers for the formula editor: which references to colour, whether the caret accepts a
// pointed reference, the function name being typed (autocomplete) and the call the caret is in
// (argument tooltip). Formula text here includes the leading '='.
import {
	FUNCTION_CATALOG,
	MAX_COL,
	MAX_ROW,
	columnIndex,
	parseRange,
	type CellRange,
	type FunctionInfo,
	referenceSpans,
} from '@christophervr/xlsx-core';

/** Excel-like colours for references in a formula, in order of first appearance. */
export const REFERENCE_COLORS = [
	'#2a6ad8',
	'#d24726',
	'#7d3fbf',
	'#0f8a3b',
	'#c0399f',
	'#946d00',
	'#d1102b',
	'#00788a',
] as const;

export const isFormulaText = (text: string): boolean =>
	text.startsWith('=') || text.startsWith('+');

export interface ColoredReference {
	/** Offsets into the formula text (which includes '='). */
	start: number;
	end: number;
	text: string;
	color: string;
	/** Same reference text shares a colour (and a colour index). */
	index: number;
}

/** The references of a formula with their colours (offsets into `text`, which includes '='). */
export function coloredReferences(text: string): ColoredReference[] {
	if (!isFormulaText(text)) return [];
	const seen = new Map<string, number>();
	return referenceSpans(text.slice(1), '').map((span) => {
		const key = span.text.replace(/\$/g, '').toUpperCase();
		let index = seen.get(key);
		if (index === undefined) {
			index = seen.size;
			seen.set(key, index);
		}
		return {
			start: span.start + 1,
			end: span.end + 1,
			text: span.text,
			index,
			color: REFERENCE_COLORS[index % REFERENCE_COLORS.length] ?? REFERENCE_COLORS[0],
		};
	});
}

/** Characters after which Excel accepts a pointed (clicked or arrowed) reference. */
const POINT_AFTER = new Set(['=', '(', ',', '+', '-', '*', '/', '^', '&', '<', '>', ':', ';', '%']);

/** Whether typing a reference at `caret` makes sense (point mode may start). */
export function acceptsReference(text: string, caret: number): boolean {
	if (!isFormulaText(text)) return false;
	if (text.slice(caret).trimStart() !== '' && !/^[\s),;&+\-*/^<>=]/.test(text.slice(caret)))
		return false;
	const before = text.slice(0, caret).trimEnd();
	const last = before[before.length - 1];
	return last !== undefined && POINT_AFTER.has(last) && !insideString(text, caret);
}

/** True when `caret` is inside a "string literal". */
export function insideString(text: string, caret: number): boolean {
	let open = false;
	for (let i = 0; i < caret && i < text.length; i++) if (text[i] === '"') open = !open;
	return open;
}

export interface NameToken {
	start: number;
	prefix: string;
}

/** The function-name prefix ending at the caret (`=SU|` -> `SU`), for autocomplete. */
export function nameAtCaret(text: string, caret: number): NameToken | undefined {
	if (!isFormulaText(text) || insideString(text, caret)) return undefined;
	const match = /[A-Za-z_][A-Za-z0-9_.]*$/.exec(text.slice(0, caret));
	if (!match) return undefined;
	const start = caret - match[0].length;
	const prev = text[start - 1] ?? '';
	if (start === 0 || /[A-Za-z0-9_$!'".:]/.test(prev)) return undefined;
	if (/^[A-Za-z0-9_.]/.test(text.slice(caret))) return undefined;
	return { start, prefix: match[0] };
}

/** Catalog functions starting with `prefix` (case-insensitive), at most `limit`. */
export function functionMatches(
	prefix: string,
	limit = 12,
	catalog: readonly FunctionInfo[] = FUNCTION_CATALOG,
): FunctionInfo[] {
	const upper = prefix.toUpperCase();
	if (!upper) return [];
	const out: FunctionInfo[] = [];
	for (const info of catalog) {
		if (info.name.startsWith(upper)) out.push(info);
		if (out.length >= limit) break;
	}
	return out;
}

export interface CallContext {
	name: string;
	/** Zero-based argument the caret is in. */
	argIndex: number;
}

/** The innermost function call around the caret (`=IF(A1>0,|` -> IF, argument 1). */
export function callAtCaret(text: string, caret: number): CallContext | undefined {
	if (!isFormulaText(text)) return undefined;
	const stack: CallContext[] = [];
	let inString = false;
	for (let i = 1; i < caret && i < text.length; i++) {
		const ch = text[i];
		if (ch === '"') inString = !inString;
		if (inString) continue;
		if (ch === '(') {
			const name = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(text.slice(0, i))?.[1];
			stack.push({ name: (name ?? '').toUpperCase(), argIndex: 0 });
		} else if (ch === ')') stack.pop();
		else if (ch === ',' || ch === ';') {
			const top = stack[stack.length - 1];
			if (top) top.argIndex++;
		} else if (ch === '{') stack.push({ name: '', argIndex: 0 });
		else if (ch === '}') stack.pop();
	}
	for (let i = stack.length - 1; i >= 0; i--) {
		const entry = stack[i];
		if (entry?.name) return entry;
		if (entry && !entry.name) return undefined;
	}
	return undefined;
}

/** Splits `SUM(number1, [number2], ...)` into its parameter texts. */
export function syntaxParams(syntax: string): { name: string; params: string[] } {
	const open = syntax.indexOf('(');
	const close = syntax.lastIndexOf(')');
	if (open < 0) return { name: syntax, params: [] };
	const inner = syntax.slice(open + 1, close < 0 ? undefined : close).trim();
	const params = inner ? inner.split(',').map((p) => p.trim()) : [];
	return { name: syntax.slice(0, open), params };
}

/** The parameter to highlight for an argument index (a trailing `...` repeats the last one). */
export function activeParam(params: string[], argIndex: number): number {
	if (argIndex < params.length) {
		if (params[argIndex] === '...') return Math.max(0, argIndex - 1);
		return argIndex;
	}
	return params[params.length - 1] === '...' ? params.length - 2 : params.length - 1;
}

/** Replaces `text[start, end)` with `insert`; returns the new text and caret. */
export function splice(
	text: string,
	start: number,
	end: number,
	insert: string,
): { text: string; caret: number } {
	return { text: text.slice(0, start) + insert + text.slice(end), caret: start + insert.length };
}

/** The sheet name and range a reference token names (`'My sheet'!$A$1:B2`, `A:C`, `2:5`). */
export function referenceTarget(text: string): { sheet?: string; range: CellRange } | undefined {
	const bang = text.lastIndexOf('!');
	const sheet =
		bang >= 0 ? text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'") : undefined;
	const body = text
		.slice(bang + 1)
		.replace(/\$/g, '')
		.toUpperCase();
	let range: CellRange | undefined;
	const cols = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(body);
	const rows = /^(\d+):(\d+)$/.exec(body);
	if (cols) {
		const a = columnIndex(cols[1] ?? 'A');
		const b = columnIndex(cols[2] ?? 'A');
		range = { start: { row: 0, col: Math.min(a, b) }, end: { row: MAX_ROW, col: Math.max(a, b) } };
	} else if (rows) {
		const a = Number(rows[1]) - 1;
		const b = Number(rows[2]) - 1;
		range = { start: { row: Math.min(a, b), col: 0 }, end: { row: Math.max(a, b), col: MAX_COL } };
	} else range = parseRange(body);
	if (!range) return undefined;
	return sheet === undefined ? { range } : { sheet, range };
}

const cycleCell = (cell: string): string => {
	const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(cell);
	if (!m) return cell;
	const [, c, col, r, row] = m;
	// A1 -> $A$1 -> A$1 -> $A1 -> A1, like Excel's F4.
	if (!c && !r) return `$${col}$${row}`;
	if (c && r) return `${col}$${row}`;
	if (!c && r) return `$${col}${row}`;
	return `${col}${row}`;
};

/** F4: cycles the absolute/relative form of the reference at the caret. */
export function toggleAbsolute(
	text: string,
	caret: number,
): { text: string; caret: number } | undefined {
	const ref = coloredReferences(text).find((r) => caret >= r.start && caret <= r.end);
	if (!ref) return undefined;
	const bang = ref.text.lastIndexOf('!');
	const prefix = ref.text.slice(0, bang + 1);
	const body = ref.text
		.slice(bang + 1)
		.split(':')
		.map(cycleCell)
		.join(':');
	const next = splice(text, ref.start, ref.end, prefix + body);
	return next;
}

/** Excel's commit-time autocorrect: appends the closing parentheses a formula is missing. */
export function closeParens(text: string): string {
	if (!isFormulaText(text)) return text;
	let depth = 0;
	let inString = false;
	for (const ch of text) {
		if (ch === '"') inString = !inString;
		else if (!inString && ch === '(') depth++;
		else if (!inString && ch === ')') depth = Math.max(0, depth - 1);
	}
	return depth > 0 && !inString ? text + ')'.repeat(depth) : text;
}
