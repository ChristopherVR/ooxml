/** SVG-compatible affine coefficients, shared by geometry and bounded vector scenes. */
export type AffineMatrix = readonly [number, number, number, number, number, number];
export const IDENTITY_AFFINE: AffineMatrix = Object.freeze([1, 0, 0, 1, 0, 0] as const);
export function composeAffine(a: AffineMatrix, b: AffineMatrix): AffineMatrix {
	return [
		a[0] * b[0] + a[2] * b[1],
		a[1] * b[0] + a[3] * b[1],
		a[0] * b[2] + a[2] * b[3],
		a[1] * b[2] + a[3] * b[3],
		a[0] * b[4] + a[2] * b[5] + a[4],
		a[1] * b[4] + a[3] * b[5] + a[5],
	];
}
