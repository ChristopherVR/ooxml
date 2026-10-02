// Find & Select > Go To Special: selects the cells of a kind (formulas, constants, blanks, notes,
// conditional formats, data validation, current region, last cell) within the selection, or the
// used range when a single cell is selected.
import {
	type CellAddress,
	type CellRange,
	currentRegion,
	getCell,
	isCellError,
	normalizeRange,
	rangesIntersect,
	usedRange,
} from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { target } from './util.js';

export type SpecialKind =
	| 'formulas'
	| 'constants'
	| 'blanks'
	| 'comments'
	| 'conditional'
	| 'validation'
	| 'currentRegion'
	| 'lastCell';

export interface SpecialQuery {
	kind: SpecialKind;
	/** For formulas / constants: which value types (all when absent). */
	types?: ReadonlyArray<'numbers' | 'text' | 'logicals' | 'errors'>;
}

const MAX_RANGES = 10_000;

/** Selects the matching cells; false (and no change) when nothing matches. */
export function selectSpecial(ctx: EditorContext, query: SpecialQuery): boolean {
	const ranges = findSpecial(ctx, query);
	if (!ranges.length) return false;
	const first = ranges[0]!;
	ctx.selection.set({ ranges, anchor: first.start, active: first.start });
	ctx.grid()?.scrollTo(first.start);
	return true;
}

export function findSpecial(ctx: EditorContext, query: SpecialQuery): CellRange[] {
	const t = target(ctx);
	if (!t) return [];
	const used = usedRange(t.ws);
	if (query.kind === 'currentRegion') return [currentRegion(t.ws, t.active)];
	if (query.kind === 'lastCell') return used ? [{ start: used.end, end: used.end }] : [];
	if (!used) return [];
	const single =
		t.ranges.length === 1 &&
		t.range.start.row === t.range.end.row &&
		t.range.start.col === t.range.end.col;
	const scopes = single ? [used] : t.ranges.map(normalizeRange);
	const typeOk = (value: unknown): boolean => {
		const types = query.types;
		if (!types?.length) return true;
		if (typeof value === 'number') return types.includes('numbers');
		if (typeof value === 'string') return types.includes('text');
		if (typeof value === 'boolean') return types.includes('logicals');
		if (isCellError(value)) return types.includes('errors');
		return false;
	};
	const at = (cell: CellAddress): CellRange => ({ start: cell, end: cell });
	const matches = (row: number, col: number): boolean => {
		const cell = getCell(t.ws, row, col);
		switch (query.kind) {
			case 'formulas':
				return cell?.formula !== undefined && typeOk(cell.value);
			case 'constants':
				return (
					!!cell &&
					cell.formula === undefined &&
					cell.value !== null &&
					cell.value !== '' &&
					typeOk(cell.value)
				);
			case 'blanks':
				return !cell || cell.value === null || cell.value === '';
			case 'comments':
				return t.ws.comments.some((c) => c.address.row === row && c.address.col === col);
			case 'conditional':
				return t.ws.conditionalFormats.some((cf) =>
					cf.ranges.some((r) => rangesIntersect(r, at({ row, col }))),
				);
			case 'validation':
				return t.ws.dataValidations.some((dv) =>
					dv.ranges.some((r) => rangesIntersect(r, at({ row, col }))),
				);
			default:
				return false;
		}
	};
	const out: CellRange[] = [];
	for (const scope of scopes) {
		const s = {
			start: {
				row: Math.max(scope.start.row, used.start.row),
				col: Math.max(scope.start.col, used.start.col),
			},
			end: {
				row: Math.min(scope.end.row, used.end.row),
				col: Math.min(scope.end.col, used.end.col),
			},
		};
		for (let row = s.start.row; row <= s.end.row; row++) {
			let runStart = -1;
			for (let col = s.start.col; col <= s.end.col + 1; col++) {
				const hit = col <= s.end.col && matches(row, col);
				if (hit && runStart < 0) runStart = col;
				if (!hit && runStart >= 0) {
					out.push({ start: { row, col: runStart }, end: { row, col: col - 1 } });
					runStart = -1;
					if (out.length >= MAX_RANGES) return out;
				}
			}
		}
	}
	return out;
}
