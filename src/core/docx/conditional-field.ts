// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// IF field (ECMA-376 Part 1, 17.16.5.25): IF expr1 operator expr2 trueText [falseText].
import { applyFieldFormat, parseFieldInstruction } from './field-format';

type ComparisonOp = '=' | '<>' | '<' | '<=' | '>' | '>=';

const OPS: ReadonlySet<string> = new Set(['=', '<>', '<', '<=', '>', '>=']);

function parseNumeric(text: string): number | undefined {
	const trimmed = text.trim();
	if (!trimmed) return undefined;
	const n = Number(trimmed);
	return Number.isFinite(n) ? n : undefined;
}

/** Word IF wildcard matching for `=` and `<>`: `*` matches any sequence, `?` matches one char. */
function wildcardMatch(text: string, pattern: string): boolean {
	const s = text.toLowerCase();
	const p = pattern.toLowerCase();
	let si = 0;
	let pi = 0;
	let star = -1;
	let match = 0;
	while (si < s.length) {
		if (pi < p.length && (p[pi] === '?' || p[pi] === s[si])) {
			si++;
			pi++;
		} else if (pi < p.length && p[pi] === '*') {
			star = pi++;
			match = si;
		} else if (star >= 0) {
			pi = star + 1;
			si = ++match;
		} else {
			return false;
		}
	}
	while (pi < p.length && p[pi] === '*') pi++;
	return pi === p.length;
}

function compare(left: string, op: ComparisonOp, right: string): boolean {
	const ln = parseNumeric(left);
	const rn = parseNumeric(right);
	if (ln !== undefined && rn !== undefined) {
		switch (op) {
			case '=':
				return ln === rn;
			case '<>':
				return ln !== rn;
			case '<':
				return ln < rn;
			case '<=':
				return ln <= rn;
			case '>':
				return ln > rn;
			case '>=':
				return ln >= rn;
		}
	}
	if (op === '=') return wildcardMatch(left, right);
	if (op === '<>') return !wildcardMatch(left, right);
	const cmp = left.toLowerCase().localeCompare(right.toLowerCase());
	switch (op) {
		case '<':
			return cmp < 0;
		case '<=':
			return cmp <= 0;
		case '>':
			return cmp > 0;
		case '>=':
			return cmp >= 0;
	}
}

/**
 * Evaluates a Word `IF` field instruction. Returns undefined when the instruction has an
 * unsupported switch, missing operands, or an unrecognized operator so the saved result is kept.
 */
export function ifFieldResult(instr: string): string | undefined {
	const field = parseFieldInstruction(instr);
	if (field?.name !== 'IF') return undefined;
	if (field.args.length < 4 || field.args.length > 5) return undefined;
	const [left, op, right, trueText, falseText] = field.args;
	if (left === undefined || op === undefined || right === undefined || trueText === undefined)
		return undefined;
	if (!OPS.has(op)) return undefined;
	const chosen = compare(left, op as ComparisonOp, right) ? trueText : (falseText ?? '');
	return applyFieldFormat(chosen, field.format);
}
