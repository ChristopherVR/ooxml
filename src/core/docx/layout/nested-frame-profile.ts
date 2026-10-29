import type { LayoutBorder, LayoutTable } from './input';
export interface NestedFrameProfile {
	leftPx: number;
	rightPx: number;
	bottomPx: number;
}
const equal = (a: LayoutBorder | undefined, b: LayoutBorder | undefined) =>
	a !== undefined &&
	b !== undefined &&
	a.style === b.style &&
	a.color === b.color &&
	Math.abs(a.widthPx - b.widthPx) < 1e-6;
/**
 * A vertical merge removes the edge it covers: the continuing cell has no top, the cell above it no
 * bottom. Those absent edges are not conflicts, so the comparisons below skip them.
 */
function mergedAway(
	table: LayoutTable,
	rowIndex: number,
	column: number,
	side: 'top' | 'bottom',
	merges: boolean,
) {
	if (!merges) return false;
	const row = table.rows[rowIndex]!;
	const cell = row.cells[column]!;
	if (side === 'top') return cell.verticalMerge === 'continue';
	const below = table.rows[rowIndex + 1]?.cells.find(
		(other) => other.xPx === cell.xPx && other.widthPx === cell.widthPx,
	);
	return below?.verticalMerge === 'continue';
}
/** Native evidence covers complete solid edges with consistent shared borders and row grids. */
export function nestedFrameProfile(
	table: LayoutTable,
	tableWidthPx = table.widthPx,
	/** Body tables only: vertical merges and rows merged differently are measured to keep the model. */
	merges = false,
): NestedFrameProfile | undefined {
	const sides = table.rows.flatMap((row, rowIndex) =>
		row.cells.flatMap((cell, column) =>
			(['top', 'right', 'bottom', 'left'] as const)
				.filter(
					(side) =>
						(side !== 'top' && side !== 'bottom') ||
						!mergedAway(table, rowIndex, column, side, merges),
				)
				.map((side) => cell.borders?.[side]),
		),
	);
	if (sides.every((side) => !side)) return { leftPx: 0, rightPx: 0, bottomPx: 0 };
	if (!sides.every((side) => side?.style === 'solid' && side.widthPx > 0)) return undefined;
	// With grid positions on every cell, rows may be merged differently (a cell over several columns).
	const gridKnown =
		merges &&
		table.rows.every((row) =>
			row.cells.every((cell) => cell.xPx !== undefined && cell.widthPx !== undefined),
		);
	const first = table.rows[0]!,
		left = first.cells[0]!.borders!.left!,
		right = first.cells.at(-1)!.borders!.right!;
	for (const [index, row] of table.rows.entries()) {
		if (
			(!gridKnown && row.cells.length !== first.cells.length) ||
			!equal(row.cells[0]?.borders?.left, left) ||
			!equal(row.cells.at(-1)?.borders?.right, right)
		)
			return undefined;
		for (const [column, cell] of row.cells.entries()) {
			const sourceWidth =
				cell.widthPx ?? (tableWidthPx === undefined ? undefined : tableWidthPx / row.cells.length);
			const paintWidth =
				sourceWidth === undefined
					? undefined
					: sourceWidth +
						(cell.borders!.left!.widthPx - cell.borders!.right!.widthPx) / 2 +
						(column === row.cells.length - 1 ? cell.borders!.right!.widthPx : 0);
			if (paintWidth !== undefined && paintWidth <= 0) return undefined;

			const topAway = mergedAway(table, index, column, 'top', merges);
			const bottomAway = mergedAway(table, index, column, 'bottom', merges);
			// Each row's edges are compared with those of its first cell that still has them.
			const reference = (side: 'top' | 'bottom') =>
				row.cells.find((_, at) => !mergedAway(table, index, at, side, merges))?.borders?.[side];
			if (
				(!topAway && !equal(cell.borders?.top, reference('top'))) ||
				(!bottomAway && !equal(cell.borders?.bottom, reference('bottom')))
			)
				return undefined;
			if (column && !equal(row.cells[column - 1]?.borders?.right, cell.borders?.left))
				return undefined;
			if (index) {
				const above = table.rows[index - 1]!;
				const previous = gridKnown
					? above.cells.find(
							(other) =>
								other.xPx! <= cell.xPx! + 1e-6 && cell.xPx! < other.xPx! + other.widthPx! - 1e-6,
						)
					: above.cells[column];
				if (
					!previous ||
					(!gridKnown && (previous.xPx !== cell.xPx || previous.widthPx !== cell.widthPx)) ||
					(!topAway &&
						!mergedAway(table, index - 1, above.cells.indexOf(previous), 'bottom', merges) &&
						!equal(previous.borders?.bottom, cell.borders?.top))
				)
					return undefined;
			}
		}
	}
	return {
		leftPx: left.widthPx,
		rightPx: right.widthPx,
		bottomPx: table.rows.at(-1)!.cells[0]!.borders!.bottom!.widthPx,
	};
}
