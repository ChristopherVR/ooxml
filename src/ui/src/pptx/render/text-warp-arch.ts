/**
 * The arch ellipse PowerPoint draws for Arch Up, Arch Down and Button: the
 * box's own ellipse, one half of it, with the text baseline on the path.
 *
 * Measured on PowerPoint 16 exports (a row of full-block glyphs, so the ink is
 * the line box): the ellipse is centred horizontally, has the box's own radii
 * and is shifted down by the default text frame inset, so the baseline of a
 * single line passes that inset below the top edge (arch up) or the bottom
 * edge (arch down) and the arc ends at mid height. The glyph tops point away
 * from the centre for Arch Up and Button and toward it for Arch Down. A
 * following line lies on a concentric ellipse: `t` 0 is the outermost, 1 the
 * innermost.
 *
 * The start angle does not turn the arc: it only picks which half carries the
 * text (see {@link archOnUpperHalf}).
 */

/** Default start angle of Arch Up and Button (180 degrees: the left tip). */
const ARCH_DEFAULT_ADJ = 10800000;

/**
 * How far the ellipse sits below the box (px at 96 dpi): the default text
 * frame inset (0.05 in) plus the baseline-to-line-box offset measured for
 * Arial, which differs a little between the two halves.
 */
const ARCH_FRAME_INSET = { up: 6.3, down: 8.5 } as const;

/** An angle in 60000ths of a degree, as degrees in [0, 360). */
export function normalizeWarpAngle(adj: number): number {
	const turn = 21600000;
	return (((adj % turn) + turn) % turn) / 60000;
}

/**
 * Whether an arch or button preset puts its text on the upper half of the
 * ellipse. Measured on PowerPoint 16 at every 15 degrees: from 90 up to, but
 * not including, 270 degrees the text lies on the upper half of the box
 * ellipse, from 270 round to 90 on the lower half. Arch Up and Button default
 * to 180 degrees (upper half), Arch Down to 0 (lower half).
 */
export function archOnUpperHalf(adj: number | undefined, half: 'up' | 'down'): boolean {
	const degrees = normalizeWarpAngle(adj ?? (half === 'up' ? ARCH_DEFAULT_ADJ : 0));
	return degrees >= 90 && degrees < 270;
}

/**
 * SVG path for the arch half-ellipse. `half` is the preset's own half (`up`:
 * glyph tops away from the centre, `down`: toward it); a start angle in the
 * other range turns the whole picture half a turn about the centre of the
 * box, so Arch Up reads upside down along the lower half.
 */
export function archEllipsePath(
	width: number,
	height: number,
	t: number,
	adj: number | undefined,
	half: 'up' | 'down',
): string {
	const scale = Math.max(0.1, 1 + (0.5 - t) * 0.7);
	const rx = (width / 2) * scale;
	const ry = (height / 2) * scale;
	const turned = archOnUpperHalf(adj, half) !== (half === 'up');
	const inset = ARCH_FRAME_INSET[half];
	const x0 = width / 2 - rx;
	const x1 = width / 2 + rx;
	const sweep = half === 'up' ? 1 : 0;
	if (!turned) {
		const cy = height / 2 + inset;
		return `M ${x0},${cy} A ${rx},${ry} 0 0,${sweep} ${x1},${cy}`;
	}
	const cy = height / 2 - inset;
	return `M ${x1},${cy} A ${rx},${ry} 0 0,${sweep} ${x0},${cy}`;
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
 * ellipse run once around clockwise, with the text centred on the path. The
 * start angle only picks the tip: from 180 round to 360 degrees (the default)
 * the loop starts at the left tip, so a centred line sits on the right tip and
 * runs down the right side; from 0 to 180 it starts at the right tip and the
 * text sits on the left. A following line lies on a concentric ellipse: `t` 0
 * is the outermost, 1 the innermost.
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
	const startsLeft = normalizeWarpAngle(adj ?? CIRCLE_DEFAULT_ADJ) >= 180;
	const degrees = (startsLeft ? 180 : 0) + CIRCLE_FIT.degrees;
	const angle = degrees * (Math.PI / 180);
	const x0 = cx + rx * Math.cos(angle);
	const y0 = cy + ry * Math.sin(angle);
	const x1 = 2 * cx - x0;
	const y1 = 2 * cy - y0;
	return `M ${x0},${y0} A ${rx},${ry} 0 0,1 ${x1},${y1} A ${rx},${ry} 0 0,1 ${x0},${y0}`;
}
