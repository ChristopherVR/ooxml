import { executableCellFormula } from './cell-formula';

const RATIO = '[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:e[-+]?\\d+)?';
const PLAIN = new RegExp(`^\\s*(Width|Height)\\s*\\*\\s*${RATIO}\\s*$`, 'i');
const BOUNDED = new RegExp(
	`^(\\s*BOUND\\(\\s*)(Width|Height)\\s*\\*\\s*${RATIO}(\\s*,[\\s\\S]*\\))\\s*$`,
	'i',
);

/** The dimension a control position is measured against: X against Width, Y against Height. */
export const controlDimension = (axis: 'x' | 'y') => (axis === 'x' ? 'Width' : 'Height');

/**
 * How Visio stores a dragged control position: a ratio of the shape size (`Width*0.3`), inside
 * the cell's own `BOUND(...)` limits when it has them. Returns the new formula, `''` for a cell
 * that holds a plain value (the value is written and the formula stays empty), or `undefined`
 * when the cell is computed some other way (`Width*0.5+User.Offset`), which a drag cannot
 * rewrite without losing that relationship.
 */
export function rewriteControlFormula(
	formula: string | undefined,
	axis: 'x' | 'y',
	ratio: number,
): string | undefined {
	if (!Number.isFinite(ratio)) return undefined;
	const source = executableCellFormula(formula);
	const dimension = controlDimension(axis);
	const number = String(Number(ratio.toPrecision(12)));
	if (source === undefined) return '';
	const plain = PLAIN.exec(source);
	if (plain)
		return plain[1]!.toLowerCase() === dimension.toLowerCase()
			? `${dimension}*${number}`
			: undefined;
	const bounded = BOUNDED.exec(source);
	if (bounded && bounded[2]!.toLowerCase() === dimension.toLowerCase())
		return `${bounded[1]}${dimension}*${number}${bounded[3]}`;
	return undefined;
}

/** Whether a drag can rewrite this formula (see `rewriteControlFormula`). */
export const controlFormulaDraggable = (formula: string | undefined, axis: 'x' | 'y'): boolean =>
	rewriteControlFormula(formula, axis, 0) !== undefined;
