import type { Cell, CellValue } from '../model.js';
import { scalarBinary } from '../formula/operators.js';
import { parseNumberText } from '../formula/text-number.js';
import type { PasteOperation } from './types.js';

const OPS = { add: '+', subtract: '-', multiply: '*', divide: '/' } as const;
const participates = (value: CellValue): boolean =>
	value === null ||
	typeof value === 'number' ||
	(typeof value === 'string' && parseNumberText(value) !== undefined);
const operand = (cell: Pick<Cell, 'value' | 'formula'>): string =>
	cell.formula !== undefined
		? `(${cell.formula})`
		: String(
				typeof cell.value === 'string'
					? parseNumberText(cell.value)
					: cell.value === null
						? 0
						: cell.value,
			);

/** Paste arithmetic ignores nonnumeric constants, but combines existing formulas. */
export function pasteArithmetic(
	destination: Cell | undefined,
	source: Pick<Cell, 'value' | 'formula'>,
	operation: Exclude<PasteOperation, 'none'>,
): Pick<Cell, 'value' | 'formula'> {
	const dest = destination ?? { value: null };
	if (
		(!source.formula && !participates(source.value)) ||
		(!dest.formula && !participates(dest.value))
	)
		return { value: dest.value, ...(dest.formula !== undefined ? { formula: dest.formula } : {}) };
	if (source.value === null && dest.value === null && !source.formula && !dest.formula)
		return { value: null };
	const op = OPS[operation];
	const result: Pick<Cell, 'value' | 'formula'> = {
		value: scalarBinary(op, dest.value, source.value),
	};
	if (dest.formula !== undefined || source.formula !== undefined)
		result.formula = `${operand(dest)}${op}${operand(source)}`;
	return result;
}
