// Rewriting references for row and column inserts and deletes and for sheet renames. Every
// formula in the workbook is rewritten on each such edit, so a formula's split into literal text
// and references (its template) is cached by text, and the template of the rewritten text is
// derived instead of tokenizing it again: the next edit reads it from the cache.
import { MAX_COL, MAX_ROW } from '../address';
import type { RefSpec, SheetPrefix } from './ast';
import { prefixText, rewriteRef, safeTokens, sameSpec } from './ref-text';
import type { Token } from './tokenizer';

export interface ShiftSpec {
	/** Sheet whose rows or columns are inserted or deleted. */
	sheet: string;
	axis: 'row' | 'col';
	/** Zero-based first row or column affected. */
	at: number;
	/** Positive to insert, negative to delete. */
	count: number;
}

/** A reference (or a sheet-qualified name) inside a template. */
interface RefPart {
	kind: 'ref' | 'name';
	text: string;
	prefix?: SheetPrefix;
	/** The prefix's sheet in lower case; undefined for external references. */
	sheet?: string;
	/** Absent for `#REF!`. */
	ref?: RefSpec;
	spill?: boolean;
}

/** A formula as literal text runs and the parts rewriters may change. */
type Template = (string | RefPart)[];

const templates = new Map<string, Template | null>();
const MAX_TEMPLATES = 200_000;

/** Empties the template cache (tests compare cached rewrites with fresh ones). */
export function forgetFormulaTemplates(): void {
	templates.clear();
}

function remember(text: string, template: Template | null): void {
	if (templates.size >= MAX_TEMPLATES) templates.clear();
	templates.set(text, template);
}

/**
 * Joins `A1 : A5` (spaces around the colon) into one area token, the way Excel reads it when
 * rows or columns shift, so deleting row 1 gives `A1:A4` rather than `#REF! : A4`.
 */
function mergeSpacedAreas(tokens: Token[]): Token[] {
	const out: Token[] = [];
	const plainCell = (t: Token | undefined): boolean =>
		t?.kind === 'ref' && t.ref?.kind === 'cell' && !t.spill;
	for (let i = 0; i < tokens.length; i++) {
		const first = tokens[i] as Token;
		let j = i + 1;
		while (tokens[j]?.kind === 'ws') j++;
		let k = j + 1;
		while (tokens[k]?.kind === 'ws') k++;
		const colon = tokens[j];
		const last = tokens[k];
		if (
			plainCell(first) &&
			colon?.kind === 'op' &&
			colon.text === ':' &&
			plainCell(last) &&
			!last?.prefix &&
			k > i + 2 &&
			first.ref &&
			last?.ref
		) {
			const text = tokens
				.slice(i, k + 1)
				.map((t) => t.text)
				.join('');
			out.push({
				...first,
				text,
				ref: { kind: 'area', start: first.ref.start, end: last.ref.start },
			});
			i = k;
			continue;
		}
		out.push(first);
	}
	return out;
}

/** The cached template of a formula; null when it cannot be tokenized. */
function templateOf(formula: string): Template | null {
	const cached = templates.get(formula);
	if (cached !== undefined) return cached;
	const raw = safeTokens(formula);
	let template: Template | null = null;
	if (raw) {
		template = [];
		let literal = '';
		for (const token of mergeSpacedAreas(raw)) {
			if (token.kind !== 'ref' && !(token.kind === 'name' && token.prefix)) {
				literal += token.text;
				continue;
			}
			if (literal) template.push(literal);
			literal = '';
			const part: RefPart = { kind: token.kind, text: token.text };
			if (token.prefix) {
				part.prefix = token.prefix;
				if (token.prefix.book === undefined) part.sheet = token.prefix.sheet.toLowerCase();
			}
			if (token.ref) part.ref = token.ref;
			if (token.spill) part.spill = true;
			template.push(part);
		}
		if (literal) template.push(literal);
	}
	remember(formula, template);
	return template;
}

/** Applies `rewrite` to a formula's parts; the result's template is cached for the next edit. */
function rewriteTemplate(formula: string, rewrite: (part: RefPart) => RefPart | undefined): string {
	const template = templateOf(formula);
	if (!template) return formula;
	let next: Template | undefined;
	for (let i = 0; i < template.length; i++) {
		const part = template[i] as string | RefPart;
		if (typeof part === 'string') continue;
		const changed = rewrite(part);
		if (!changed) continue;
		next ??= template.slice();
		next[i] = changed;
	}
	if (!next) return formula;
	const text = next.map((part) => (typeof part === 'string' ? part : part.text)).join('');
	remember(text, next);
	return text;
}

/** Shifts one coordinate interval for an insert or delete; `undefined` when it is deleted entirely. */
function shiftInterval(
	a: number,
	b: number,
	at: number,
	count: number,
	max: number,
): [number, number] | undefined {
	if (count > 0) {
		const na = a >= at ? a + count : a;
		const nb = b >= at ? b + count : b;
		if (nb > max) return na > max ? undefined : [na, max];
		return [na, nb];
	}
	const n = -count;
	const last = at + n - 1;
	if (a >= at && b <= last) return undefined;
	const na = a < at ? a : a <= last ? at : a - n;
	const nb = b < at ? b : b <= last ? at - 1 : b - n;
	return [na, nb];
}

/** A reference moved for the shift: the same spec when it does not move, undefined when deleted. */
function shiftRef(ref: RefSpec, spec: ShiftSpec): RefSpec | undefined {
	const isRow = spec.axis === 'row';
	if ((isRow && ref.kind === 'cols') || (!isRow && ref.kind === 'rows')) return ref;
	const startFirst = isRow ? ref.start.row <= ref.end.row : ref.start.col <= ref.end.col;
	const lo = startFirst ? ref.start : ref.end;
	const hi = startFirst ? ref.end : ref.start;
	const shifted = shiftInterval(
		isRow ? lo.row : lo.col,
		isRow ? hi.row : hi.col,
		spec.at,
		spec.count,
		isRow ? MAX_ROW : MAX_COL,
	);
	if (!shifted) return undefined;
	const [a, b] = shifted;
	const nlo = isRow ? { ...lo, row: a } : { ...lo, col: a };
	const nhi = isRow ? { ...hi, row: b } : { ...hi, col: b };
	const next: RefSpec = {
		kind: ref.kind,
		start: startFirst ? nlo : nhi,
		end: startFirst ? nhi : nlo,
	};
	if (ref.kind === 'cell') next.end = next.start;
	return sameSpec(next, ref) ? ref : next;
}

/**
 * Adjusts references for inserted (`count` > 0) or deleted (`count` < 0) rows or columns on
 * `spec.sheet`. References into deleted cells become `#REF!`.
 */
export function shiftFormula(formula: string, formulaSheet: string, spec: ShiftSpec): string {
	if (spec.count === 0) return formula;
	const sheet = spec.sheet.toLowerCase();
	const own = formulaSheet.toLowerCase() === sheet;
	return rewriteTemplate(formula, (part) => {
		if (part.kind !== 'ref' || !part.ref || (part.prefix ? part.sheet !== sheet : !own))
			return undefined;
		const next = shiftRef(part.ref, spec);
		if (next === part.ref) return undefined;
		const out: RefPart = { kind: 'ref', text: rewriteRef(part, next) };
		if (part.prefix) out.prefix = part.prefix;
		if (part.sheet !== undefined) out.sheet = part.sheet;
		if (next) out.ref = next;
		if (part.spill) out.spill = true;
		return out;
	});
}

/** Rewrites sheet prefixes naming `oldName` (case-insensitive) to `newName`, quoting as needed. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
	const lower = oldName.toLowerCase();
	return rewriteTemplate(formula, (part) => {
		const prefix = part.prefix;
		if (!prefix || part.sheet === undefined) return undefined;
		const first = part.sheet === lower;
		const second = prefix.sheet2?.toLowerCase() === lower;
		if (!first && !second) return undefined;
		const renamed: SheetPrefix = { ...prefix, sheet: first ? newName : prefix.sheet };
		if (prefix.sheet2 !== undefined) renamed.sheet2 = second ? newName : prefix.sheet2;
		const text = prefixText(renamed);
		renamed.text = text;
		const out: RefPart = { ...part, text: text + part.text.slice(prefix.text.length) };
		out.prefix = renamed;
		out.sheet = renamed.sheet.toLowerCase();
		return out;
	});
}
