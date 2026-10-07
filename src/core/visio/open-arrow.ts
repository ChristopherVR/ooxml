/** Unit span in inches of open BeginArrow/EndArrow codes 1, 3 and 9 at unit drawing scale.
 * Measured against Visio 16 SVG exports: seven sizes and three line weights.
 * The stroke contributes to the span; it is not a stroke-relative multiplier.
 */
export function visioOpenArrowExtent(size: number, lineWidth: number): number {
	const index = Number.isFinite(size) ? Math.max(0, Math.min(6, Math.trunc(size))) : 2;
	const base = [0.02, 0.025, 0.035, 0.045, 0.055, 0.125, 0.25][index]!;
	return base + (Number.isFinite(lineWidth) ? Math.max(0, lineWidth) : 0);
}
