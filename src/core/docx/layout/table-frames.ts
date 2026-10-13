import type { LayoutTableBox } from './result';
/** Clip decorative frames to a physical cell fragment without introducing another text story. */
export function clipTableFrames(
	tables: LayoutTableBox[],
	start: number,
	end: number,
	inset = 0,
): LayoutTableBox[] {
	return tables.flatMap((table) => {
		const top = Math.max(start, table.yPx);
		const bottom = Math.min(end, table.yPx + table.heightPx);
		if (bottom <= top) return [];
		const rows = table.rows.flatMap((row) => {
			const rowTop = table.yPx + row.yPx;
			const a = Math.max(top, rowTop),
				b = Math.min(bottom, rowTop + row.heightPx);
			if (b <= a) return [];
			return [
				{
					...row,
					yPx: a - top,
					heightPx: b - a,
					...(row.cellTables
						? {
								cellTables: row.cellTables.map((nested) =>
									clipTableFrames(nested, a - rowTop, b - rowTop),
								),
							}
						: {}),
				},
			];
		});
		if (table.fragmentEdges && rows.length) {
			const edges = table.fragmentEdges;
			for (const [row, side] of [
				[rows[0]!, 'top'],
				[rows.at(-1)!, 'bottom'],
			] as const) {
				if (row.geometry)
					row.geometry = row.geometry.map((cell, index) => ({
						...cell,
						borders: { ...cell.borders, [side]: edges[side][index]! },
					}));
			}
		}
		return [{ ...table, yPx: top - start + inset, heightPx: bottom - top, rows }];
	});
}
