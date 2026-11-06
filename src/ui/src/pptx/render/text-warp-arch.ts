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

/** Default start angle of Circle (180 degrees: the left tip). */
const CIRCLE_DEFAULT_ADJ = 10800000;

/**
 * The loop's offsets from the box ellipse (px at 96 dpi) and the angle the
 * text is centred off the tip. They are a least-squares fit of a text-on-path
 * ellipse to three PowerPoint 16 captures of Arial 32 pt (the loop is flatter
 * than the box and rotated slightly, as PowerPoint bends the glyphs along it
 * rather than rotating them rigidly), not derived from the preset's geometry.
 */
const CIRCLE_FIT = { dx: 2.25, dy: 1, drx: 3, dry: -12.75, degrees: -9.5 } as const;

/**
 * The closed loop PowerPoint draws for Circle: an ellipse near the box
 * ellipse from the start angle `adj` (default 180 degrees, the left tip) once
 * around clockwise, with the text centred on the path, so at the default
 * angle it sits on the right tip and runs down the right side. A following
 * line lies on a concentric ellipse: `t` 0 is the outermost, 1 the innermost.
 */
export function circleEllipsePath(
	width: number,
	height: number,
	t: number,
	adj: number | undefined,
): string {
	const scale = Math.max(0.1, 1 + (0.5 - t) * 0.7);
	const rx = (width / 2) * scale + CIRCLE_FIT.drx;
	const ry = (height / 2) * scale + CIRCLE_FIT.dry;
	const cx = width / 2 + CIRCLE_FIT.dx;
	const cy = height / 2 + CIRCLE_FIT.dy;
	const degrees = ((adj ?? CIRCLE_DEFAULT_ADJ) % 21600000) / 60000 + CIRCLE_FIT.degrees;
	const angle = degrees * (Math.PI / 180);
	const x0 = cx + rx * Math.cos(angle);
	const y0 = cy + ry * Math.sin(angle);
	const x1 = 2 * cx - x0;
	const y1 = 2 * cy - y0;
	return `M ${x0},${y0} A ${rx},${ry} 0 0,1 ${x1},${y1} A ${rx},${ry} 0 0,1 ${x0},${y0}`;
}
