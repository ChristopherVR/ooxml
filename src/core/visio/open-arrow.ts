/** Unit span in inches of open BeginArrow/EndArrow codes 1, 3, 7 and 9 at unit drawing scale.
 * Measured against Visio 16 SVG exports: seven sizes and three line weights.
 * The stroke contributes to the span; it is not a stroke-relative multiplier.
 */
export function visioOpenArrowExtent(size: number, lineWidth: number): number {
	const index = Number.isFinite(size) ? Math.max(0, Math.min(6, Math.trunc(size))) : 2;
	const base = [0.02, 0.025, 0.035, 0.045, 0.055, 0.125, 0.25][index]!;
	return base + (Number.isFinite(lineWidth) ? Math.max(0, lineWidth) : 0);
}

// Numerical geometry measured from native Visio 16 exports, in native y-down
// coordinates. Core converts it to local y-up inches; no SVG markup is imported.
const nativeGlyphs: Readonly<Record<number, string>> = {
	1: 'M 1 -1 L 0 0 L 1 1',
	3: 'M 2 1 L 0 0 L 2 -1',
	7: 'M 1.84309 -0.959456 C 1.43792 -0.380801 0.784503 -0.0260255 0.078527 -0.00137091 L 0.0785229 0.00137043 C 0.784499 0.0260231 1.43791 0.380798 1.84309 0.959452',
	9: 'M 1 -1 L -1 1',
};

/** Supported native open-arrow path in local y-up inches, anchored at (0, 0).
 * Returns undefined for styles requiring a different glyph or endpoint setback.
 */
export function visioOpenArrowPath(
	code: number,
	size: number,
	lineWidth: number,
): string | undefined {
	const glyph = nativeGlyphs[code];
	if (!glyph) return undefined;
	const extent = visioOpenArrowExtent(size, lineWidth);
	let index = 0;
	return glyph.replace(/-?\d+(?:\.\d+)?/g, (value) =>
		String(Number(value) * extent * (index++ % 2 ? 1 : -1)),
	);
}
