// What a row or column insert or delete does to the formula graph, decided without parsing the
// rewritten formulas again: how each precedent area moves, and which formulas can show a different
// value afterwards (see engine-structure.ts).
import { type CellRange, MAX_COL, MAX_ROW } from '../address';
import type { FormulaAst } from './ast';
import { getFunction } from './functions/registry';
import type { Area } from './values';

/** A structural edit on one sheet: `count` rows or columns inserted (positive) or deleted at `at`. */
export interface GraphShift {
	axis: 'row' | 'col';
	at: number;
	count: number;
}

/** What a formula's syntax says about how a structural edit can change its value. */
export interface AstTraits {
	/**
	 * Reads areas the formula text does not spell out (defined names, tables, 3D references):
	 * their precedents are worked out again when an edit reaches them.
	 */
	complex: boolean;
	/** Calls a name (a LAMBDA) whose cells the graph does not know: recalculated on every edit. */
	opaque: boolean;
	/** Returns a position (ROW, COLUMN), so moving the formula or its precedents changes it. */
	positional: boolean;
	/** Returns text about references or sheets (FORMULATEXT, CELL, SHEET). */
	descriptive: boolean;
	/** Sheet names its prefixes mention, lower case. */
	sheets: ReadonlySet<string>;
}

const POSITIONAL = new Set(['ROW', 'COLUMN']);
const DESCRIPTIVE = new Set(['FORMULATEXT', 'CELL', 'SHEET']);
const traitsCache = new WeakMap<FormulaAst, AstTraits>();

/** The traits of a parsed formula (cached per syntax tree). */
export function astTraits(ast: FormulaAst): AstTraits {
	const cached = traitsCache.get(ast);
	if (cached) return cached;
	const sheets = new Set<string>();
	const traits = { complex: false, opaque: false, positional: false, descriptive: false, sheets };
	const stack: FormulaAst[] = [ast];
	while (stack.length) {
		const node = stack.pop() as FormulaAst;
		switch (node.type) {
			case 'ref':
				if (node.prefix) {
					sheets.add(node.prefix.sheet.toLowerCase());
					if (node.prefix.sheet2 !== undefined) {
						sheets.add(node.prefix.sheet2.toLowerCase());
						traits.complex = true;
					}
				}
				break;
			case 'name':
				if (node.prefix) sheets.add(node.prefix.sheet.toLowerCase());
				traits.complex = true;
				break;
			case 'structured':
				traits.complex = true;
				break;
			case 'unary':
			case 'percent':
				stack.push(node.operand);
				break;
			case 'binary':
				stack.push(node.left, node.right);
				break;
			case 'call':
				if (POSITIONAL.has(node.name)) traits.positional = true;
				if (DESCRIPTIVE.has(node.name)) traits.descriptive = true;
				if (!getFunction(node.name)) traits.complex = traits.opaque = true;
				stack.push(...node.args);
				break;
			case 'invoke':
				stack.push(node.callee, ...node.args);
				break;
			default:
				break;
		}
	}
	traitsCache.set(ast, traits);
	return traits;
}

/** How an interval moves: by `delta`, or `undefined` when its cells change (cut, widened, gone). */
export function intervalDelta(a: number, b: number, shift: GraphShift): number | undefined {
	const { at, count } = shift;
	if (count > 0) {
		if (b < at) return 0;
		if (a < at) return undefined;
		return b + count > (shift.axis === 'row' ? MAX_ROW : MAX_COL) ? undefined : count;
	}
	const last = at - count - 1;
	if (b < at) return 0;
	if (a > last) return count;
	return undefined;
}

/** How an area moves when `shift` applies to `sheet` (areas on other sheets stay). */
export function areaDelta(area: Area, sheet: number, shift: GraphShift): number | undefined {
	if (area.sheet !== sheet) return 0;
	const { start, end } = area.range;
	return shift.axis === 'row'
		? intervalDelta(start.row, end.row, shift)
		: intervalDelta(start.col, end.col, shift);
}

/** A copy of a range moved `delta` along the axis. */
export function movedRange(range: CellRange, axis: GraphShift['axis'], delta: number): CellRange {
	if (delta === 0) return range;
	return axis === 'row'
		? {
				start: { row: range.start.row + delta, col: range.start.col },
				end: { row: range.end.row + delta, col: range.end.col },
			}
		: {
				start: { row: range.start.row, col: range.start.col + delta },
				end: { row: range.end.row, col: range.end.col + delta },
			};
}

/** Whether any of the areas reaches the moved part of the sheet (at or past the edit). */
export function reachesEdit(areas: readonly Area[], sheet: number, shift: GraphShift): boolean {
	for (const area of areas) {
		if (area.sheet !== sheet) continue;
		if ((shift.axis === 'row' ? area.range.end.row : area.range.end.col) >= shift.at) return true;
	}
	return false;
}

/**
 * Whether implicit intersection with an area can pick a different cell: the area spans several
 * positions along the axis, moves differently from the formula, and the formula's position
 * (`pos` before, `pos + nodeDelta` after) falls inside it before or after.
 */
export function intersectionMoves(
	area: Area,
	delta: number,
	pos: number,
	nodeDelta: number,
	axis: GraphShift['axis'],
): boolean {
	if (delta === nodeDelta) return false;
	const { start, end } = area.range;
	const a = axis === 'row' ? start.row : start.col;
	const b = axis === 'row' ? end.row : end.col;
	if (a === b) return false;
	return (pos >= a && pos <= b) || (pos + nodeDelta >= a + delta && pos + nodeDelta <= b + delta);
}
