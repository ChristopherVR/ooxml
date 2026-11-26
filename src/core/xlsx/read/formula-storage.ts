// Dynamic-array syntax as SpreadsheetML stores it: the spill operator `A1#` is written as
// `ANCHORARRAY(A1)` and the implicit-intersection operator `@x` as `SINGLE(x)` (both with the
// `_xlfn.` prefix the caller adds or strips).
import { FormulaError } from '../formula/ast';
import { joinTokens, type Token, tokenize } from '../formula/tokenizer';

function tokens(formula: string): Token[] | undefined {
	try {
		return tokenize(formula);
	} catch (e) {
		if (e instanceof FormulaError) return undefined;
		throw e;
	}
}

const nextIndex = (list: readonly Token[], from: number): number => {
	let i = from;
	while (list[i]?.kind === 'ws') i++;
	return i;
};

/** Index of the `close` token matching the `open` token at `open`, or -1. */
function matching(list: readonly Token[], open: number): number {
	let depth = 0;
	for (let i = open; i < list.length; i++) {
		const kind = list[i]?.kind;
		if (kind === 'open') depth++;
		else if (kind === 'close' && --depth === 0) return i;
	}
	return -1;
}

function storeField(reference: Token, field: Token): void {
	const source = reference.spill ? `ANCHORARRAY(${reference.text.slice(0, -1)})` : reference.text;
	const name = String(field.value).replaceAll('"', '""');
	reference.text = `_FV(${source},"${name}")`;
	if (reference.spill) reference.spill = false;
	field.text = '';
	field.kind = 'ws';
}

/** Writes reference operators in SpreadsheetML syntax, including `A1.Field` as `_FV`. */
export function toStoredSyntax(formula: string): string {
	if (!/[#@.]/.test(formula)) return formula;
	const list = tokens(formula);
	if (!list) return formula;
	let changed = false;
	for (let i = 0; i < list.length; i++) {
		const token = list[i] as Token;
		const field = list[i + 1];
		if (field?.kind === 'field') {
			storeField(token, field);
			changed = true;
			continue;
		}
		if (token.kind === 'ref' && token.spill) {
			token.text = `ANCHORARRAY(${token.text.slice(0, -1)})`;
			changed = true;
			continue;
		}
		if (token.kind !== 'op' || token.text !== '@') continue;
		const j = nextIndex(list, i + 1);
		const target = list[j];
		if (!target) continue;
		const targetField = list[j + 1];
		if (targetField?.kind === 'field') storeField(target, targetField);
		let end = j;
		if (target.kind === 'func') {
			end = matching(list, nextIndex(list, j + 1));
			if (end < 0) continue;
		} else if (target.kind !== 'ref' && target.kind !== 'name' && target.kind !== 'structured')
			continue;
		token.text = 'SINGLE(';
		const last = list[end] as Token;
		last.text = `${last.kind === 'ref' && last.spill ? `ANCHORARRAY(${last.text.slice(0, -1)})` : last.text})`;
		if (last.kind === 'ref') last.spill = false;
		changed = true;
	}
	return changed ? joinTokens(list) : formula;
}

/** `ANCHORARRAY(A1)` -> `A1#` and `SINGLE(ref)` -> `@ref` (after prefixes were stripped). */
export function fromStoredSyntax(formula: string): string {
	if (!/ANCHORARRAY|SINGLE/i.test(formula)) return formula;
	const list = tokens(formula);
	if (!list) return formula;
	let changed = false;
	for (let i = 0; i < list.length; i++) {
		const token = list[i] as Token;
		if (token.kind !== 'func') continue;
		const name = String(token.value ?? '').toUpperCase();
		if (name !== 'ANCHORARRAY' && name !== 'SINGLE') continue;
		const open = nextIndex(list, i + 1);
		const arg = nextIndex(list, open + 1);
		const close = nextIndex(list, arg + 1);
		const target = list[arg];
		if (list[open]?.kind !== 'open' || list[close]?.kind !== 'close' || !target) continue;
		const single = target.kind === 'ref' || target.kind === 'name' || target.kind === 'structured';
		if (name === 'ANCHORARRAY' && !(target.kind === 'ref' && target.ref?.kind === 'cell')) continue;
		if (name === 'SINGLE' && !single) continue;
		const text = name === 'ANCHORARRAY' ? `${target.text}#` : `@${target.text}`;
		for (let k = i; k <= close; k++) (list[k] as Token).text = '';
		token.text = text;
		changed = true;
	}
	return changed ? joinTokens(list) : formula;
}

const PARAMETER_PREFIX = /^_xlpm\./i;
const plainName = (token: Token): string =>
	String(token.value ?? token.text)
		.replace(PARAMETER_PREFIX, '')
		.toUpperCase();

/** The token ranges of a call's arguments, from just inside its `(` to its `)`, split at commas. */
function argumentsOf(list: readonly Token[], open: number, close: number): [number, number][] {
	const args: [number, number][] = [];
	let depth = 0;
	let start = open + 1;
	for (let i = open + 1; i < close; i++) {
		const kind = list[i]?.kind;
		if (kind === 'open' || kind === 'lbrace') depth++;
		else if (kind === 'close' || kind === 'rbrace') depth--;
		else if (kind === 'comma' && depth === 0) {
			args.push([start, i - 1]);
			start = i + 1;
		}
	}
	args.push([start, close - 1]);
	return args;
}

/** The one name an argument consists of (whitespace aside), or -1. */
function soleName(list: readonly Token[], [from, to]: [number, number]): number {
	let found = -1;
	for (let i = from; i <= to; i++) {
		const kind = list[i]?.kind;
		if (kind === 'ws') continue;
		if (kind !== 'name' || found >= 0) return -1;
		found = i;
	}
	return found;
}

/**
 * Writes the names LET and LAMBDA declare with the `_xlpm.` prefix Excel stores them with,
 * where they are declared and wherever they are used in scope (`LET(x,2,x*3)` is stored as
 * `LET(_xlpm.x,2,_xlpm.x*3)`); Excel refuses a file whose parameters lack it.
 */
export function toStoredParameters(formula: string): string {
	if (!/\b(LET|LAMBDA)\s*\(/i.test(formula)) return formula;
	const list = tokens(formula);
	if (!list) return formula;
	// Each declared name with the tokens it covers: its own and those of its scope.
	const scopes: { name: string; from: number; to: number }[] = [];
	const declared = new Set<number>();
	for (let i = 0; i < list.length; i++) {
		const token = list[i] as Token;
		if (token.kind !== 'func') continue;
		const fn = String(token.value ?? '').toUpperCase();
		if (fn !== 'LET' && fn !== 'LAMBDA') continue;
		const open = nextIndex(list, i + 1);
		if (list[open]?.kind !== 'open') continue;
		const close = matching(list, open);
		if (close < 0) continue;
		const args = argumentsOf(list, open, close);
		const last = args.length - 1;
		for (let a = 0; a < last; a += fn === 'LET' ? 2 : 1) {
			const at = soleName(list, args[a] as [number, number]);
			if (at < 0) continue;
			declared.add(at);
			const from =
				fn === 'LET' ? (args[a + 1] as [number, number])[0] : (args[last] as [number, number])[0];
			scopes.push({ name: plainName(list[at] as Token), from, to: close });
		}
	}
	if (!declared.size) return formula;
	let changed = false;
	for (let i = 0; i < list.length; i++) {
		const token = list[i] as Token;
		if ((token.kind !== 'name' && token.kind !== 'func') || PARAMETER_PREFIX.test(token.text))
			continue;
		const name = plainName(token);
		if (!declared.has(i) && !scopes.some((s) => s.name === name && i >= s.from && i <= s.to))
			continue;
		token.text = `_xlpm.${token.text}`;
		changed = true;
	}
	return changed ? joinTokens(list) : formula;
}

/** Removes the `_xlpm.` prefix from LET and LAMBDA names, as Excel shows them. */
export function fromStoredParameters(formula: string): string {
	if (!/_xlpm\./i.test(formula)) return formula;
	const list = tokens(formula);
	if (!list) return formula;
	for (const token of list)
		if ((token.kind === 'name' || token.kind === 'func') && PARAMETER_PREFIX.test(token.text))
			token.text = token.text.replace(PARAMETER_PREFIX, '');
	return joinTokens(list);
}
