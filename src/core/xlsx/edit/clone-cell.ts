import type { Cell } from '../model';

/**
 * A deep copy of a cell, equal to `structuredClone(cell)`. A cell's fields are plain data, mostly
 * primitives, so they are copied directly and only nested values (rich text runs, an array
 * range, an error value) go through `structuredClone`: copying the many cells of a bulk edit
 * then costs a spread each instead of a structured clone each.
 */
export function cloneCell(cell: Cell): Cell {
	const copy = { ...cell };
	const fields = copy as Record<string, unknown>;
	for (const key of Object.keys(fields)) {
		const value = fields[key];
		if (typeof value === 'object' && value !== null) fields[key] = structuredClone(value);
	}
	return copy;
}
