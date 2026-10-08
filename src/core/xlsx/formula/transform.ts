// Token-level formula rewriting: everything except the touched references keeps the user's text.
import { type CellRange, MAX_COL, MAX_ROW, normalizeRange } from '../address';
import type { RefCorner } from './ast';
import { rewriteRef, safeTokens, sameSpec } from './ref-text';
import { joinTokens } from './tokenizer';

export { formatRefSpec, prefixText, safeTokens } from './ref-text';
export { renameSheetInFormula, type ShiftSpec, shiftFormula } from './shift-formula';

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
		token.text = rewriteRef(token, next);
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
