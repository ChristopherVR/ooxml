// Writing a formula's result into the sheet: plain values, CSE array ranges, or spills.
import { type CellAddress, MAX_COL, MAX_ROW } from '../address.js';
import type { Cell, Worksheet } from '../model.js';
import { sameRange } from './engine-util.js';
import type { FormulaNode } from './graph.js';
import {
	clearFootprint,
	fillArrayRange,
	footprintFor,
	footprintFree,
	writeFootprint,
} from './spill.js';
import { ERR, Matrix, type Scalar } from './values.js';

/**
 * Stores `result` for `node`: a CSE formula fills its array range, an array result spills into
 * free neighbouring cells (or shows #SPILL!), anything else becomes the cell value.
 */
export function storeResult(
	sheet: Worksheet,
	node: FormulaNode,
	cell: Cell,
	result: Scalar | Matrix,
	claims: (anchor: CellAddress, row: number, col: number) => boolean,
): void {
	const { row, col } = node;
	if (node.arrayRange) {
		fillArrayRange(
			sheet,
			node.arrayRange,
			result instanceof Matrix ? result : new Matrix([[result]]),
		);
		return;
	}
	if (!(result instanceof Matrix)) {
		if (node.spill) clearFootprint(sheet, node.spill, row, col);
		delete node.spill;
		delete node.blockedSpill;
		cell.value = result;
		return;
	}
	const footprint = footprintFor(row, col, result);
	if (!footprint || !footprintFree(sheet, footprint, row, col, claims)) {
		if (node.spill) clearFootprint(sheet, node.spill, row, col);
		delete node.spill;
		node.blockedSpill = footprint ?? {
			start: { row, col },
			end: {
				row: Math.min(MAX_ROW, row + result.rows - 1),
				col: Math.min(MAX_COL, col + result.cols - 1),
			},
		};
		cell.value = ERR.SPILL;
		return;
	}
	// Own spill cells are valid destinations. Reuse them when the shape stays the same,
	// but still check for user overwrites, merges and competing spills on every pass.
	if (node.spill && !sameRange(node.spill, footprint)) clearFootprint(sheet, node.spill, row, col);
	delete node.blockedSpill;
	writeFootprint(sheet, footprint, result);
	node.spill = footprint;
}
