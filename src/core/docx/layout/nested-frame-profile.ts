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
/** Native evidence covers complete solid edges with consistent shared borders and row grids. */
export function nestedFrameProfile(
	table: LayoutTable,
	tableWidthPx = table.widthPx,
): NestedFrameProfile | undefined {
	const sides = table.rows.flatMap((row) =>
		row.cells.flatMap((cell) =>
			(['top', 'right', 'bottom', 'left'] as const).map((side) => cell.borders?.[side]),
		),
	);
	if (sides.every((side) => !side)) return { leftPx: 0, rightPx: 0, bottomPx: 0 };
	if (!sides.every((side) => side?.style === 'solid' && side.widthPx > 0)) return undefined;
	const first = table.rows[0]!,
		left = first.cells[0]!.borders!.left!,
		right = first.cells.at(-1)!.borders!.right!;
	for (const [index, row] of table.rows.entries()) {
		if (
			row.cells.length !== first.cells.length ||
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

			if (
				!equal(cell.borders?.top, row.cells[0]?.borders?.top) ||
				!equal(cell.borders?.bottom, row.cells[0]?.borders?.bottom)
			)
				return undefined;
			if (column && !equal(row.cells[column - 1]?.borders?.right, cell.borders?.left))
				return undefined;
			if (index) {
				const previous = table.rows[index - 1]!.cells[column]!;
				if (
					previous.xPx !== cell.xPx ||
					previous.widthPx !== cell.widthPx ||
					!equal(previous.borders?.bottom, cell.borders?.top)
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
