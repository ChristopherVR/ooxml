/** Tile alignment extracted from the PowerPoint image painter. Fractions anchor the same tile and frame points. */
export function pictureTileAlignment(alignment?: string): { x: number; y: number } {
	const values: Record<string, [number, number]> = {
		tl: [0, 0],
		t: [0.5, 0],
		tr: [1, 0],
		l: [0, 0.5],
		ctr: [0.5, 0.5],
		r: [1, 0.5],
		bl: [0, 1],
		b: [0.5, 1],
		br: [1, 1],
	};
	const [x, y] = values[alignment ?? 'tl'] ?? [0, 0];
	return { x, y };
}
export function pictureTileGeometry(
	frame: { width: number; height: number },
	source: { width: number; height: number },
	tile: {
		scaleX?: number;
		scaleY?: number;
		offsetXEmu?: number;
		offsetYEmu?: number;
		alignment?: string;
		flip?: string;
	},
	dpi = 96,
):
	| { width: number; height: number; x: number; y: number; columns: number; rows: number }
	| undefined {
	const width = (source.width * (tile.scaleX ?? 1) * 96) / dpi;
	const height = (source.height * (tile.scaleY ?? 1) * 96) / dpi;
	if (![width, height].every((n) => Number.isFinite(n) && n > 0)) return undefined;
	const anchor = pictureTileAlignment(tile.alignment);
	return {
		width,
		height,
		x: (frame.width - width) * anchor.x + (tile.offsetXEmu ?? 0) / 9525,
		y: (frame.height - height) * anchor.y + (tile.offsetYEmu ?? 0) / 9525,
		columns: tile.flip === 'x' || tile.flip === 'xy' ? 2 : 1,
		rows: tile.flip === 'y' || tile.flip === 'xy' ? 2 : 1,
	};
}
/** Inverse of the host's rotation-then-flip transform, centered on its frame. */
export function pictureFillCounterTransform(
	width: number,
	height: number,
	rotation = 0,
	flipH = false,
	flipV = false,
): string | undefined {
	if (!rotation && !flipH && !flipV) return undefined;
	return `translate(${width / 2} ${height / 2}) scale(${flipH ? -1 : 1} ${flipV ? -1 : 1}) rotate(${-rotation}) translate(${-width / 2} ${-height / 2})`;
}
