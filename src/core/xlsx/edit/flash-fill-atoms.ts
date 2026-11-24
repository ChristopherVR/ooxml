/** Case transformation applied to an extracted substring. */
export type CaseMode = 'raw' | 'upper' | 'lower' | 'proper';

type DelimKind = 'ws' | ',' | '-' | '_' | '/' | '\\' | '.' | '@' | ':' | ';' | '|' | '(' | ')';
type RegexKind = 'digits' | 'number' | 'letters' | 'alnum';
type DatePartKind = 'year' | 'month' | 'monthNum' | 'day' | 'dayNum' | 'monthShort' | 'monthLong';

export type AtomOp =
	| { kind: 'whole' }
	| { kind: 'token'; delim: DelimKind; index: number }
	| { kind: 'tokenPrefix'; delim: DelimKind; index: number; len: number }
	| { kind: 'regex'; regex: RegexKind; index: number }
	| { kind: 'between'; left: string; right: string }
	| { kind: 'prefix'; len: number }
	| { kind: 'suffix'; len: number }
	| { kind: 'slice'; start: number; len: number }
	| { kind: 'date'; part: DatePartKind };

export interface AtomExpr {
	col: number;
	op: AtomOp;
	casing: CaseMode;
	cost: number;
}

export interface FlashFillExample {
	inputs: readonly string[];
	output: string;
}

const DELIMS: readonly DelimKind[] = ['ws', ',', '-', '_', '/', '\\', '.', '@', ':', ';', '|'];
const BETWEEN_PAIRS: readonly (readonly [string, string])[] = [
	['(', ')'],
	['[', ']'],
	['<', '>'],
	['@', '.'],
	['-', '-'],
	['/', '/'],
	['_', '_'],
	[' ', ' '],
];
const CASES: readonly CaseMode[] = ['raw', 'upper', 'lower', 'proper'];
const REGEXES: Record<RegexKind, RegExp> = {
	digits: /\d+/g,
	number: /\d+(?:\.\d+)?/g,
	letters: /[A-Za-z]+/g,
	alnum: /[A-Za-z0-9]+/g,
};

const MONTH_SHORT = [
	'Jan',
	'Feb',
	'Mar',
	'Apr',
	'May',
	'Jun',
	'Jul',
	'Aug',
	'Sep',
	'Oct',
	'Nov',
	'Dec',
] as const;
const MONTH_LONG = [
	'January',
	'February',
	'March',
	'April',
	'May',
	'June',
	'July',
	'August',
	'September',
	'October',
	'November',
	'December',
] as const;
const MONTH_MAP = new Map<string, number>([
	...MONTH_SHORT.map((m, i): [string, number] => [m.toLowerCase(), i + 1]),
	...MONTH_LONG.map((m, i): [string, number] => [m.toLowerCase(), i + 1]),
]);

function applyCase(text: string, casing: CaseMode): string {
	switch (casing) {
		case 'raw':
			return text;
		case 'upper':
			return text.toUpperCase();
		case 'lower':
			return text.toLowerCase();
		case 'proper':
			return text.replace(
				/[A-Za-z]+/g,
				(word) => (word[0]?.toUpperCase() ?? '') + word.slice(1).toLowerCase(),
			);
	}
}

function splitDelim(text: string, delim: DelimKind): string[] {
	return delim === 'ws'
		? text.trim().split(/\s+/).filter(Boolean)
		: text.split(delim).map((part) => part.trim());
}

function pickIndex(items: readonly string[], index: number): string | undefined {
	const pos = index < 0 ? items.length + index : index;
	return pos >= 0 && pos < items.length ? items[pos] : undefined;
}

interface ParsedDate {
	year: number;
	month: number;
	day: number;
}

const ISO_DATE_RE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\b|$)/;
const SLASH_DATE_RE = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/;
const TEXT_DATE_RE = /^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/;
const DAY_MONTH_RE = /^(\d{1,2})[-\s]([A-Za-z]+)[-\s](\d{2,4})$/;

const expandYear = (y: number): number => (y < 100 ? (y >= 30 ? 1900 + y : 2000 + y) : y);

function parseDateText(raw: string): ParsedDate | undefined {
	const text = raw.trim();
	let m = ISO_DATE_RE.exec(text);
	if (m) {
		const [year, month, day] = [Number(m[1]), Number(m[2]), Number(m[3])];
		if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return { year, month, day };
	}
	m = SLASH_DATE_RE.exec(text);
	if (m) {
		const [a, b, year] = [Number(m[1]), Number(m[2]), expandYear(Number(m[3]))];
		if (a >= 1 && a <= 12 && b >= 1 && b <= 31) return { year, month: a, day: b };
		if (b >= 1 && b <= 12 && a >= 1 && a <= 31) return { year, month: b, day: a };
	}
	m = TEXT_DATE_RE.exec(text);
	if (m) {
		const month = MONTH_MAP.get(m[1]!.toLowerCase());
		const [day, year] = [Number(m[2]), Number(m[3])];
		if (month && day >= 1 && day <= 31) return { year, month, day };
	}
	m = DAY_MONTH_RE.exec(text);
	if (m) {
		const [day, year] = [Number(m[1]), expandYear(Number(m[3]))];
		const month = MONTH_MAP.get(m[2]!.toLowerCase());
		if (month && day >= 1 && day <= 31) return { year, month, day };
	}
	return undefined;
}

function formatDatePart(d: ParsedDate, part: DatePartKind): string | undefined {
	switch (part) {
		case 'year':
			return String(d.year).padStart(4, '0');
		case 'month':
			return String(d.month).padStart(2, '0');
		case 'monthNum':
			return String(d.month);
		case 'day':
			return String(d.day).padStart(2, '0');
		case 'dayNum':
			return String(d.day);
		case 'monthShort':
			return MONTH_SHORT[d.month - 1];
		case 'monthLong':
			return MONTH_LONG[d.month - 1];
	}
}

export function evalAtom(expr: AtomExpr, inputs: readonly string[]): string | undefined {
	const src = inputs[expr.col];
	if (src === undefined) return undefined;
	const { op } = expr;
	let raw: string | undefined;
	switch (op.kind) {
		case 'whole':
			raw = src;
			break;
		case 'token':
			raw = pickIndex(splitDelim(src, op.delim), op.index);
			break;
		case 'tokenPrefix': {
			const tok = pickIndex(splitDelim(src, op.delim), op.index);
			raw = tok !== undefined && tok.length >= op.len ? tok.slice(0, op.len) : undefined;
			break;
		}
		case 'regex': {
			const matches = src.match(REGEXES[op.regex]);
			raw = matches ? pickIndex(matches, op.index) : undefined;
			break;
		}
		case 'between': {
			const start = src.indexOf(op.left);
			if (start < 0) return undefined;
			const after = start + op.left.length;
			const end = src.indexOf(op.right, after);
			raw = end >= 0 ? src.slice(after, end) : undefined;
			break;
		}
		case 'prefix':
			raw = src.length >= op.len ? src.slice(0, op.len) : undefined;
			break;
		case 'suffix':
			raw = src.length >= op.len ? src.slice(src.length - op.len) : undefined;
			break;
		case 'slice':
			raw = src.length >= op.start + op.len ? src.slice(op.start, op.start + op.len) : undefined;
			break;
		case 'date': {
			const d = parseDateText(src);
			raw = d ? formatDatePart(d, op.part) : undefined;
			break;
		}
	}
	return raw !== undefined && raw.length > 0 ? applyCase(raw, expr.casing) : undefined;
}

export function generateAtomsForExamples(
	first: FlashFillExample,
	all: readonly FlashFillExample[],
): AtomExpr[] {
	const out: AtomExpr[] = [];
	const push = (col: number, op: AtomOp, baseCost: number): void => {
		for (const casing of CASES) {
			const expr: AtomExpr = { col, op, casing, cost: baseCost + (casing === 'raw' ? 0 : 0.4) };
			const v0 = evalAtom(expr, first.inputs);
			if (!v0 || !first.output.includes(v0)) continue;
			let ok = true;
			for (let i = 1; i < all.length; i++) {
				const vm = evalAtom(expr, all[i]!.inputs);
				if (!vm || !all[i]!.output.includes(vm)) {
					ok = false;
					break;
				}
			}
			if (ok) out.push(expr);
		}
	};

	for (let col = 0; col < first.inputs.length; col++) {
		const src = first.inputs[col] ?? '';
		if (!src) continue;
		push(col, { kind: 'whole' }, 1);

		for (const delim of DELIMS) {
			const parts = splitDelim(src, delim);
			if (parts.length < 2) continue;
			for (let idx = 0; idx < parts.length; idx++) {
				push(col, { kind: 'token', delim, index: idx }, 2);
				push(col, { kind: 'token', delim, index: idx - parts.length }, 2.1);
				const part = parts[idx] ?? '';
				const maxPre = Math.min(3, part.length - 1);
				for (let len = 1; len <= maxPre; len++) {
					push(col, { kind: 'tokenPrefix', delim, index: idx, len }, 2.6);
					push(col, { kind: 'tokenPrefix', delim, index: idx - parts.length, len }, 2.7);
				}
			}
		}

		for (const regex of ['digits', 'number', 'letters', 'alnum'] as const) {
			const matches = src.match(REGEXES[regex]);
			if (!matches?.length) continue;
			for (let idx = 0; idx < matches.length; idx++) {
				push(col, { kind: 'regex', regex, index: idx }, 2.8);
				push(col, { kind: 'regex', regex, index: idx - matches.length }, 2.9);
			}
		}

		for (const [left, right] of BETWEEN_PAIRS) push(col, { kind: 'between', left, right }, 2.6);

		if (parseDateText(src)) {
			for (const part of [
				'year',
				'month',
				'monthNum',
				'day',
				'dayNum',
				'monthShort',
				'monthLong',
			] as const)
				push(col, { kind: 'date', part }, 2.5);
		}

		const maxFix = Math.min(src.length - 1, first.output.length);
		for (let len = 1; len <= maxFix; len++) {
			push(col, { kind: 'prefix', len }, 3.0);
			push(col, { kind: 'suffix', len }, 3.1);
		}
		if (src.length <= 48) {
			for (let start = 1; start < src.length - 1; start++)
				for (let len = 2; start + len < src.length; len++)
					push(col, { kind: 'slice', start, len }, 6.5);
		}
	}
	return out;
}
