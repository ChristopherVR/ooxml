// Dynamic-array syntax as SpreadsheetML stores it: the spill operator `A1#` is written as
// `ANCHORARRAY(A1)` and the implicit-intersection operator `@x` as `SINGLE(x)` (both with the
// `_xlfn.` prefix the caller adds or strips).
import { FormulaError } from '../formula/ast.js';
import { joinTokens, type Token, tokenize } from '../formula/tokenizer.js';

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

/** `A1#` -> `ANCHORARRAY(A1)`, `@x` -> `SINGLE(x)`; anything else is unchanged. */
export function toStoredSyntax(formula: string): string {
	if (!formula.includes('#') && !formula.includes('@')) return formula;
	const list = tokens(formula);
	if (!list) return formula;
	let changed = false;
	for (let i = 0; i < list.length; i++) {
		const token = list[i] as Token;
		if (token.kind === 'ref' && token.spill) {
			token.text = `ANCHORARRAY(${token.text.slice(0, -1)})`;
			changed = true;
			continue;
		}
		if (token.kind !== 'op' || token.text !== '@') continue;
		const j = nextIndex(list, i + 1);
		const target = list[j];
		if (!target) continue;
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
