/**
 * The arch ellipse PowerPoint draws for Arch Up, Arch Down and Button at the
 * default angle (10800000, 180 degrees): the box's own ellipse, one half of
 * it, with the text baseline on the path.
 *
 * Measured on PowerPoint 16 exports (a row of full-block glyphs, so the ink is
 * the line box): the ellipse is centred horizontally, has the box's own radii
 * and is shifted down by the default text frame inset, so the baseline of a
 * single line passes that inset below the top edge (arch up) or the bottom
 * edge (arch down) and the arc ends at mid height. The glyph tops point away
 * from the centre for Arch Up and Button and toward it for Arch Down. A
 * following line lies on a concentric ellipse: `t` 0 is the outermost, 1 the
 * innermost. Other angles keep the previous approximation.
 */

/** Default start angle of the arch and button presets (180 degrees). */
const ARCH_DEFAULT_ADJ = 10800000;

/**
 * How far the ellipse sits below the box (px at 96 dpi): the default text
 * frame inset (0.05 in) plus the baseline-to-line-box offset measured for
 * Arial, which differs a little between the two halves.
 */
const ARCH_FRAME_INSET = { up: 6.3, down: 8.5 } as const;

/** SVG path for the arch half-ellipse, or `undefined` when `adj` is not the default angle. */
export function archEllipsePath(
	width: number,
	height: number,
	t: number,
	adj: number | undefined,
	half: 'up' | 'down',
): string | undefined {
	if (adj !== undefined && adj !== ARCH_DEFAULT_ADJ) {
		return undefined;
	}
	const scale = Math.max(0.1, 1 + (0.5 - t) * 0.7);
	const rx = (width / 2) * scale;
	const ry = (height / 2) * scale;
	const cy = height / 2 + ARCH_FRAME_INSET[half];
	const x0 = width / 2 - rx;
	const x1 = width / 2 + rx;
	return `M ${x0},${cy} A ${rx},${ry} 0 0,${half === 'up' ? 1 : 0} ${x1},${cy}`;
}
