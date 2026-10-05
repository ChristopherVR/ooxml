// Reference rewrites for moving cells (cut and paste) and deleting sheets.
import { type CellRange, normalizeRange } from '../address.js';
import type { RefSpec } from './ast.js';
import { joinTokens } from './tokenizer.js';
import { formatRefSpec, prefixText, safeTokens } from './transform.js';

/** Cells moved by a cut and paste: `range` on `fromSheet` lands at (+dRow, +dCol) on `toSheet`. */
export interface MoveSpec {
	fromSheet: string;
	range: CellRange;
	toSheet: string;
	dRow: number;
	dCol: number;
}

const same = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

function within(spec: RefSpec, range: CellRange): boolean {
	if (spec.kind !== 'cell' && spec.kind !== 'area') return false;
	const r = normalizeRange({
		start: { row: spec.start.row, col: spec.start.col },
		end: { row: spec.end.row, col: spec.end.col },
	});
	const box = normalizeRange(range);
	return (
		r.start.row >= box.start.row &&
		r.end.row <= box.end.row &&
		r.start.col >= box.start.col &&
		r.end.col <= box.end.col
	);
}

const shifted = (box: CellRange, dRow: number, dCol: number): CellRange => ({
	start: { row: box.start.row + dRow, col: box.start.col + dCol },
	end: { row: box.end.row + dRow, col: box.end.col + dCol },
});

function withPrefix(sheet: string | undefined, body: string): string {
	return sheet === undefined ? body : prefixText({ text: '', sheet }) + body;
}

/**
 * Rewrites a formula for a cut and paste, as Excel does: references wholly inside the moved
 * range follow it (absolute parts too, onto the destination sheet), references to cells the
 * paste overwrote become `#REF!`, and other references keep pointing where they did. When the
 * formula itself moves to another sheet (`newFormulaSheet`), its unqualified references gain
 * the original sheet's name.
 */
export function moveReferencesInFormula(
	formula: string,
	formulaSheet: string,
	move: MoveSpec,
	newFormulaSheet = formulaSheet,
): string {
	const tokens = safeTokens(formula);
	if (!tokens) return formula;
	const dest = shifted(normalizeRange(move.range), move.dRow, move.dCol);
	let changed = false;
	for (const token of tokens) {
		if (token.kind !== 'ref' || token.prefix?.book !== undefined || token.prefix?.sheet2) continue;
		const target = token.prefix?.sheet ?? formulaSheet;
		const spec = token.ref;
		const spill = token.spill ? '#' : '';
		if (spec && same(target, move.fromSheet) && within(spec, move.range)) {
			const moveCorner = (c: RefSpec['start']) => ({
				...c,
				row: c.row + move.dRow,
				col: c.col + move.dCol,
			});
			const next: RefSpec = {
				kind: spec.kind,
				start: moveCorner(spec.start),
				end: moveCorner(spec.end),
			};
			const qualified = token.prefix !== undefined || !same(move.toSheet, newFormulaSheet);
			token.text = withPrefix(qualified ? move.toSheet : undefined, formatRefSpec(next) + spill);
			changed = true;
			continue;
		}
		if (spec && same(target, move.toSheet) && within(spec, dest)) {
			token.text = '#REF!';
			changed = true;
			continue;
		}
		if (!token.prefix && !same(formulaSheet, newFormulaSheet)) {
			token.text = withPrefix(formulaSheet, token.text);
			changed = true;
		}
	}
	return changed ? joinTokens(tokens) : formula;
}

/**
 * Rewrites a formula after sheet `sheetName` was deleted: references to it (and to the given
 * tables that lived on it) become `#REF!`. 3D references (`Sheet1:Sheet3!A1`) are left alone.
 */
export function deleteSheetInFormula(
	formula: string,
	sheetName: string,
	tableNames: readonly string[] = [],
): string {
	const tokens = safeTokens(formula);
	if (!tokens) return formula;
	const tables = new Set(tableNames.map((t) => t.toLowerCase()));
	let changed = false;
	for (const token of tokens) {
		const prefix = token.prefix;
		const onSheet =
			prefix !== undefined &&
			prefix.book === undefined &&
			prefix.sheet2 === undefined &&
			same(prefix.sheet, sheetName);
		const table =
			typeof token.value === 'string' &&
			tables.has(token.value.toLowerCase()) &&
			(token.kind === 'structured' || (token.kind === 'name' && !prefix));
		if (((token.kind === 'ref' || token.kind === 'name') && onSheet) || table) {
			token.text = '#REF!';
			changed = true;
		}
	}
	return changed ? joinTokens(tokens) : formula;
}
