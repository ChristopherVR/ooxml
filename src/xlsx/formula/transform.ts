// Token-level formula rewriting: everything except the touched references keeps the user's text.
import { type CellRange, columnLabel, MAX_COL, MAX_ROW, normalizeRange } from '../address.js';
import { FormulaError, type RefCorner, type RefSpec, type SheetPrefix } from './ast.js';
import { joinTokens, type Token, tokenize } from './tokenizer.js';

export interface ShiftSpec {
	/** Sheet whose rows or columns are inserted or deleted. */
	sheet: string;
	axis: 'row' | 'col';
	/** Zero-based first row or column affected. */
	at: number;
	/** Positive to insert, negative to delete. */
	count: number;
}

export function safeTokens(formula: string): Token[] | undefined {
	try {
		return tokenize(formula);
	} catch (e) {
		if (e instanceof FormulaError) return undefined;
		throw e;
	}
}

const cornerText = (c: RefCorner, kind: RefSpec['kind']): string => {
	if (kind === 'cols') return `${c.colAbs ? '$' : ''}${columnLabel(c.col)}`;
	if (kind === 'rows') return `${c.rowAbs ? '$' : ''}${c.row + 1}`;
	return `${c.colAbs ? '$' : ''}${columnLabel(c.col)}${c.rowAbs ? '$' : ''}${c.row + 1}`;
};

/** Formats a reference spec (without sheet prefix). */
export function formatRefSpec(spec: RefSpec): string {
	const a = cornerText(spec.start, spec.kind);
	if (spec.kind === 'cell') return a;
	return `${a}:${cornerText(spec.end, spec.kind)}`;
}

function rewrite(token: Token, ref: RefSpec | undefined): string {
	const prefix = token.prefix?.text ?? '';
	if (!ref) return `${prefix}#REF!`;
	return `${prefix}${formatRefSpec(ref)}${token.spill ? '#' : ''}`;
}

const sameSpec = (a: RefSpec, b: RefSpec): boolean =>
	a.kind === b.kind &&
	a.start.row === b.start.row &&
	a.start.col === b.start.col &&
	a.end.row === b.end.row &&
	a.end.col === b.end.col;

/** Moves relative references by (dRow, dCol), as copying or filling a formula does. Off-grid parts become #REF!. */
export function translateFormula(formula: string, dRow: number, dCol: number): string {
	const tokens = safeTokens(formula);
	if (!tokens) return formula;
	for (const token of tokens) {
		if (token.kind !== 'ref' || !token.ref) continue;
		const spec = token.ref;
		const move = (c: RefCorner): RefCorner | undefined => {
			const row = spec.kind === 'cols' || c.rowAbs ? c.row : c.row + dRow;
			const col = spec.kind === 'rows' || c.colAbs ? c.col : c.col + dCol;
			if (row < 0 || row > MAX_ROW || col < 0 || col > MAX_COL) return undefined;
			return { ...c, row, col };
		};
		const start = move(spec.start);
		const end = move(spec.end);
		const next = start && end ? { kind: spec.kind, start, end } : undefined;
		if (next && sameSpec(next, spec)) continue;
		token.text = rewrite(token, next);
	}
	return joinTokens(tokens);
}

const sheetMatches = (
	prefix: SheetPrefix | undefined,
	formulaSheet: string,
	sheet: string,
): boolean => {
	if (prefix?.book !== undefined) return false;
	const name = prefix ? prefix.sheet : formulaSheet;
	return name.toLowerCase() === sheet.toLowerCase();
};

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

/**
 * Adjusts references for inserted (`count` > 0) or deleted (`count` < 0) rows or columns on
 * `spec.sheet`. References into deleted cells become `#REF!`.
 */
export function shiftFormula(formula: string, formulaSheet: string, spec: ShiftSpec): string {
	if (spec.count === 0) return formula;
	const raw = safeTokens(formula);
	if (!raw) return formula;
	const tokens = mergeSpacedAreas(raw);
	const isRow = spec.axis === 'row';
	for (const token of tokens) {
		if (token.kind !== 'ref' || !token.ref || !sheetMatches(token.prefix, formulaSheet, spec.sheet))
			continue;
		const ref = token.ref;
		if ((isRow && ref.kind === 'cols') || (!isRow && ref.kind === 'rows')) continue;
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
		let next: RefSpec | undefined;
		if (shifted) {
			const [a, b] = shifted;
			const nlo = isRow ? { ...lo, row: a } : { ...lo, col: a };
			const nhi = isRow ? { ...hi, row: b } : { ...hi, col: b };
			next = { kind: ref.kind, start: startFirst ? nlo : nhi, end: startFirst ? nhi : nlo };
			if (ref.kind === 'cell') next.end = next.start;
		}
		if (next && sameSpec(next, ref)) continue;
		token.text = rewrite(token, next);
	}
	return joinTokens(tokens);
}

/**
 * Whether a sheet name must be quoted in a reference: anything but a plain identifier, and names
 * Excel would read as something else: A1 or R1C1 references (`R`, `C`, `RC`, `R1`, `C2`,
 * `R1C1`) and the logicals TRUE and FALSE.
 */
const needsQuote = (name: string): boolean =>
	!/^[A-Za-z_\u00A1-\uFFFF][A-Za-z0-9_.\u00A1-\uFFFF]*$/.test(name) ||
	/^[A-Za-z]{1,3}\d+$/.test(name) ||
	/^(?:R\d*C?\d*|C\d*)$/i.test(name) ||
	/^(?:TRUE|FALSE)$/i.test(name);

/** The text of a sheet prefix (`Sheet1!`, `'My Sheet'!`), quoted when needed. */
export function prefixText(prefix: SheetPrefix): string {
	const sheets = prefix.sheet2 !== undefined ? `${prefix.sheet}:${prefix.sheet2}` : prefix.sheet;
	const book = prefix.book !== undefined ? `[${prefix.book}]` : '';
	const quote =
		needsQuote(prefix.sheet) || (prefix.sheet2 !== undefined && needsQuote(prefix.sheet2));
	return quote ? `'${(book + sheets).replace(/'/g, "''")}'!` : `${book}${sheets}!`;
}

/** Rewrites sheet prefixes naming `oldName` (case-insensitive) to `newName`, quoting as needed. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
	const tokens = safeTokens(formula);
	if (!tokens) return formula;
	const lower = oldName.toLowerCase();
	for (const token of tokens) {
		const prefix = token.prefix;
		if (!prefix || prefix.book !== undefined) continue;
		const first = prefix.sheet.toLowerCase() === lower;
		const second = prefix.sheet2?.toLowerCase() === lower;
		if (!first && !second) continue;
		const renamed: SheetPrefix = { ...prefix, sheet: first ? newName : prefix.sheet };
		if (prefix.sheet2 !== undefined) renamed.sheet2 = second ? newName : prefix.sheet2;
		const text = prefixText(renamed);
		token.text = text + token.text.slice(prefix.text.length);
		renamed.text = text;
		token.prefix = renamed;
	}
	return joinTokens(tokens);
}

/** The ranges a formula references directly (for coloured reference highlighting). */
export function referencedRanges(
	formula: string,
	formulaSheet: string,
): { sheet: string; range: CellRange }[] {
	const tokens = safeTokens(formula);
	if (!tokens) return [];
	const out: { sheet: string; range: CellRange }[] = [];
	for (const token of tokens) {
		if (token.kind !== 'ref' || !token.ref || token.prefix?.book !== undefined) continue;
		const { start, end } = token.ref;
		out.push({
			sheet: token.prefix?.sheet ?? formulaSheet,
			range: normalizeRange({
				start: { row: start.row, col: start.col },
				end: { row: end.row, col: end.col },
			}),
		});
	}
	return out;
}

/** A reference token's position in formula text, for colouring references while editing. */
export interface ReferenceSpan {
	/** Character offsets into the formula exactly as passed (a leading `=` counts). */
	start: number;
	end: number;
	/** The reference text including its sheet prefix (`'My Sheet'!A1:B2`). */
	text: string;
	/** The sheet it points at (`formulaSheet` when unqualified). */
	sheet: string;
	range: CellRange;
}

/**
 * Every A1 reference in a formula with its character span (cells, areas, whole rows and
 * columns, sheet-qualified ones). `#REF!`, 3D and external-workbook references are skipped;
 * an unparsable formula yields no spans.
 */
export function referenceSpans(formula: string, formulaSheet: string): ReferenceSpan[] {
	const tokens = safeTokens(formula);
	if (!tokens) return [];
	const out: ReferenceSpan[] = [];
	for (const token of tokens) {
		if (token.kind !== 'ref' || !token.ref) continue;
		const prefix = token.prefix;
		if (prefix && (prefix.book !== undefined || prefix.sheet2 !== undefined)) continue;
		const { start, end } = token.ref;
		out.push({
			start: token.start,
			end: token.start + token.text.length,
			text: token.text,
			sheet: prefix?.sheet ?? formulaSheet,
			range: normalizeRange({
				start: { row: start.row, col: start.col },
				end: { row: end.row, col: end.col },
			}),
		});
	}
	return out;
}
