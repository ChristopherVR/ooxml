import type { LayoutTable } from './input';
/** Native evidence currently covers uniform single-line nested table borders only. */
export function nestedFrameBorderWidth(table: LayoutTable): number | undefined {
	const sides = table.rows.flatMap((row) =>
		row.cells.flatMap((cell) =>
			['top', 'right', 'bottom', 'left'].map(
				(side) => cell.borders?.[side as 'top' | 'right' | 'bottom' | 'left'],
			),
		),
	);
	if (sides.every((side) => !side)) return 0;
	const width = sides[0]?.widthPx;
	return width !== undefined &&
		sides.every((side) => side?.style === 'solid' && Math.abs(side.widthPx - width) < 1e-6)
		? width
		: undefined;
}
