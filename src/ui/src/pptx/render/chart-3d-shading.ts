/**
 * chart-3d-shading.ts: PowerPoint's flat chart-space shading of a 3D chart's
 * box faces. The per-face table lives in `ooxml-core/chart` (the shared 3-D
 * painter uses it too); this module re-exports it and keeps the normal-based
 * shading of the `c:shape` bar families the WebGL scene needs.
 *
 * @module chart-3d-shading
 */
export { buildChart3DBoxFaceColors, shadeChart3DFace, type Chart3DFace } from 'ooxml-core/chart';

type ShadeFamily = 'box' | 'cylinder' | 'cone' | 'pyramid';

/**
 * Per-shape shade coefficients `[x, y, z]` for {@link chart3DNormalShade}.
 * `box` reproduces the face multipliers above (right 0.64, top 0.75, front
 * 1.0); the others are fitted to `gt/chart-07..09.webp` (accent1 columns):
 * a cylinder runs 1.07x at its centre line to ~0.64x at the silhouette, a
 * cone 1.16x to ~0.7x, a pyramid's front face 1.03x and its side 0.68x. Tops
 * (cylinder cap, box top) sample 0.75x throughout.
 */
const NORMAL_SHADE: Record<ShadeFamily, readonly [number, number, number]> = {
	box: [0.64, 0.75, 1],
	cylinder: [0.62, 0.75, 1.07],
	cone: [0.7, 0.75, 1.16],
	pyramid: [0.68, 0.75, 1.03],
};

/** Multiplier for a downward-facing surface (a bar's underside), rarely visible. */
const UNDERSIDE_SHADE = 0.5;

function shadeFamily(shape: string | undefined): ShadeFamily {
	switch (shape) {
		case 'cylinder':
			return 'cylinder';
		case 'cone':
		case 'coneToMax':
			return 'cone';
		case 'pyramid':
		case 'pyramidToMax':
			return 'pyramid';
		default:
			return 'box';
	}
}

/**
 * PowerPoint's flat chart-space shade for a surface of a `c:shape` bar with
 * unit world normal `(nx, ny, nz)` (x right, y up, z toward the viewer), as
 * a multiplier of the base colour: the quadratic form
 * `cx nx^2 + cy ny^2 + cz nz^2`, which for a box's axis-aligned faces is
 * exactly the per-face table above.
 */
export function chart3DNormalShade(
	shape: string | undefined,
	nx: number,
	ny: number,
	nz: number,
): number {
	const [cx, cy, cz] = NORMAL_SHADE[shadeFamily(shape)];
	return cx * nx * nx + (ny >= 0 ? cy : UNDERSIDE_SHADE) * ny * ny + cz * nz * nz;
}
