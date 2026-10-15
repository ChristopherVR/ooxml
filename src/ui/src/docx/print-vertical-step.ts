import type { PrintEditingIndex } from './print-editing-index';

const CELL = '.dve-print-cell';
const ROW = '.dve-print-row';

function rowSibling(row: Element, direction: number): Element | undefined {
	let next = direction < 0 ? row.previousElementSibling : row.nextElementSibling;
	while (next && !next.matches(ROW))
		next = direction < 0 ? next.previousElementSibling : next.nextElementSibling;
	while (next && (next as HTMLElement).dataset.repeated === 'true')
		next = direction < 0 ? next.previousElementSibling : next.nextElementSibling;
	return next ?? undefined;
}

function cellLines(index: PrintEditingIndex, cell: Element): HTMLElement[] {
	return [...cell.querySelectorAll<HTMLElement>('.dve-print-line[data-block-id]')].filter((line) =>
		index.has(line),
	);
}

/** The line in the cell of `row` that sits under `x`, entered from above (down) or below (up). */
function lineInRow(
	index: PrintEditingIndex,
	row: Element,
	x: number,
	direction: number,
): HTMLElement | undefined {
	let best: { cell: Element; distance: number } | undefined;
	for (const cell of row.querySelectorAll(`:scope > ${CELL}`)) {
		if (!cellLines(index, cell).length) continue;
		const rect = cell.getBoundingClientRect();
		const distance = Math.max(rect.left - x, x - rect.right, 0);
		if (!best || distance < best.distance) best = { cell, distance };
	}
	if (!best) return undefined;
	const lines = cellLines(index, best.cell);
	return direction < 0 ? lines.at(-1) : lines[0];
}

/**
 * Within a cell that holds nested tables the lines of side-by-side nested cells interleave in
 * reading order, so a vertical step picks the nearest line that lies fully above or below, and
 * under `x` among equals.
 */
function nestedStep(
	index: PrintEditingIndex,
	cell: Element,
	line: HTMLElement,
	direction: number,
	x: number,
): HTMLElement | undefined {
	if (!cell.querySelector('[data-nested="true"]')) return undefined;
	const from = line.getBoundingClientRect();
	let best: { line: HTMLElement; dy: number; dx: number } | undefined;
	for (const candidate of cellLines(index, cell)) {
		if (candidate === line || candidate.closest('[data-repeated="true"]')) continue;
		const rect = candidate.getBoundingClientRect();
		const dy = direction < 0 ? from.top - rect.bottom : rect.top - from.bottom;
		if (dy < -1) continue;
		const dx = Math.max(rect.left - x, x - rect.right, 0);
		if (!best || dy < best.dy - 1 || (Math.abs(dy - best.dy) <= 1 && dx < best.dx))
			best = { line: candidate, dy, dx };
	}
	return best?.line;
}

/**
 * The line one step up or down from `line`. Outside tables that is the next rendered line; inside a
 * table it stays in the cell, then moves to the cell of the next row under the same x, and leaves
 * the table past its last or first row. Entering a table lands in the cell under `x`.
 */
export function stepVertical(
	index: PrintEditingIndex,
	line: HTMLElement,
	direction: number,
	x: number,
): HTMLElement {
	const cell = line.closest(CELL);
	let target = index.adjacent(line, direction);
	if (cell) {
		const inside = nestedStep(index, cell, line, direction, x);
		if (inside) return inside;
		if (target !== line && target.closest(CELL) === cell) return target;
		const row = cell.closest(ROW);
		const sibling = row && rowSibling(row, direction);
		const found = sibling && lineInRow(index, sibling, x, direction);
		if (found) return found;
		const table = cell.closest('.dve-print-table');
		const lines = table
			? [...table.querySelectorAll<HTMLElement>('.dve-print-line[data-block-id]')].filter(
					(candidate) => index.has(candidate) && !candidate.closest('[data-repeated="true"]'),
				)
			: [line];
		target = index.adjacent((direction < 0 ? lines[0] : lines.at(-1)) ?? line, direction);
	}
	for (
		let next = index.adjacent(target, direction);
		target.closest('[data-repeated="true"]') && next !== target;
		next = index.adjacent(target, direction)
	)
		target = next;
	const row = target.closest(ROW);
	if (row && row !== line.closest(ROW) && !row.contains(line))
		return lineInRow(index, row, x, direction) ?? target;
	return target;
}
