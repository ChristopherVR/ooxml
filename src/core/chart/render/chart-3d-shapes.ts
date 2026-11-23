// 3-D bar shape meshes (`c:shape`, ECMA-376 ST_Shape): box, cylinder, cone, coneToMax, pyramid
// and pyramidToMax. Every non-box shape is a lofted frustum along the value axis (Y for columns,
// X for horizontal bars) between a base ring and a top ring, with outward normals for backface
// culling and smooth or flat shading.
import { shade } from '../style-palette';
import {
	boxMesh,
	faceKind,
	type Chart3DMesh,
	type Chart3DMeshFace,
	type Chart3DMeshMeta,
	type Vec3,
} from './chart-3d-mesh';
import type { Chart3DBarShape } from './summary';

/** Number of facets approximating a circular cross-section (cylinder / cone). */
const CIRCLE_SIDES = 16;
const HEX = /^#[0-9a-f]{6}$/i;
const LIGHT: Vec3 = (() => {
	const v: Vec3 = [-0.35, 0.65, -0.67];
	const len = Math.hypot(...v);
	return [v[0] / len, v[1] / len, v[2] / len];
})();

export interface BarShapeBounds {
	/** Base along the value axis (y0 for columns, x0 for horizontal bars). */
	v0: number;
	/** Tip along the value axis (y1 for columns, x1 for horizontal bars). */
	v1: number;
	/** Cross-section bounds along the category axis. */
	c0: number;
	c1: number;
	/** Cross-section bounds along the depth axis. */
	z0: number;
	z1: number;
	horizontal: boolean;
	/** Value-axis zero baseline and extreme (used by `coneToMax` and `pyramidToMax`). */
	base: number;
	extreme: number;
}

const SQUARE_RING: readonly (readonly [number, number])[] = [
	[-1, -1],
	[1, -1],
	[1, 1],
	[-1, 1],
];

const CIRCLE_RING: readonly (readonly [number, number])[] = Array.from(
	{ length: CIRCLE_SIDES },
	(_, i) => {
		const a = (i / CIRCLE_SIDES) * Math.PI * 2;
		return [Math.cos(a), Math.sin(a)] as const;
	},
);

function outwardNormal(pts: readonly Vec3[], center: Vec3): Vec3 {
	const [a, b, c] = pts as [Vec3, Vec3, Vec3];
	const e1: Vec3 = [b[0] - a[0], b[1] - a[1], b[2] - a[2]];
	const e2: Vec3 = [c[0] - a[0], c[1] - a[1], c[2] - a[2]];
	let nx = e1[1] * e2[2] - e1[2] * e2[1];
	let ny = e1[2] * e2[0] - e1[0] * e2[2];
	let nz = e1[0] * e2[1] - e1[1] * e2[0];
	const len = Math.hypot(nx, ny, nz);
	if (len < 1e-9) return [0, 1, 0];
	nx /= len;
	ny /= len;
	nz /= len;
	let fx = 0;
	let fy = 0;
	let fz = 0;
	for (const p of pts) {
		fx += p[0];
		fy += p[1];
		fz += p[2];
	}
	fx = fx / pts.length - center[0];
	fy = fy / pts.length - center[1];
	fz = fz / pts.length - center[2];
	return nx * fx + ny * fy + nz * fz < 0 ? [-nx, -ny, -nz] : [nx, ny, nz];
}

function curvedShade(color: string, normal: Vec3): string {
	if (!HEX.test(color)) return color;
	const dot = normal[0] * LIGHT[0] + normal[1] * LIGHT[1] + normal[2] * LIGHT[2];
	const lit = Math.max(0, Math.min(1, 0.5 + 0.5 * dot));
	return shade(color, 0.45 * (1 - lit));
}

/** Builds a 3-D bar mesh for any `ST_Shape` value. Returns undefined for degenerate bounds. */
export function barShapeMesh(
	shape: Chart3DBarShape | undefined,
	bounds: BarShapeBounds,
	meta: Chart3DMeshMeta,
): Chart3DMesh | undefined {
	if (!shape || shape === 'box') {
		return bounds.horizontal
			? boxMesh(bounds.v0, bounds.v1, bounds.c0, bounds.c1, bounds.z0, bounds.z1, meta)
			: boxMesh(bounds.c0, bounds.c1, bounds.v0, bounds.v1, bounds.z0, bounds.z1, meta);
	}
	const { v0, v1, c0, c1, z0, z1, horizontal } = bounds;
	const rc = Math.abs(c1 - c0) / 2;
	const rz = Math.abs(z1 - z0) / 2;
	if (Math.abs(v1 - v0) < 1e-6 || rc < 1e-6 || rz < 1e-6) return undefined;
	const cc = (c0 + c1) / 2;
	const cz = (z0 + z1) / 2;
	const center: Vec3 = horizontal ? [(v0 + v1) / 2, cc, cz] : [cc, (v0 + v1) / 2, cz];
	const curved = shape === 'cylinder' || shape === 'cone' || shape === 'coneToMax';
	const ring = curved ? CIRCLE_RING : SQUARE_RING;
	const span = Math.abs(bounds.extreme - bounds.base);
	const taper = (v: number): number =>
		span < 1e-6 ? 0 : Math.max(0, Math.min(1, 1 - Math.abs(v - bounds.base) / span));
	let s0 = 1;
	let s1 = 1;
	if (shape === 'cone' || shape === 'pyramid') {
		s0 = 1;
		s1 = 0;
	} else if (shape === 'coneToMax' || shape === 'pyramidToMax') {
		s0 = taper(v0);
		s1 = taper(v1);
	}
	const vertex = (u: number, w: number, v: number, s: number): Vec3 =>
		horizontal ? [v, cc + u * rc * s, cz + w * rz * s] : [cc + u * rc * s, v, cz + w * rz * s];
	const baseRing = ring.map(([u, w]) => vertex(u, w, v0, s0));
	const topRing = ring.map(([u, w]) => vertex(u, w, v1, s1));
	const apexBase = vertex(0, 0, v0, 0);
	const apexTop = vertex(0, 0, v1, 0);
	const faces: Chart3DMeshFace[] = [];
	const dir = v1 >= v0 ? 1 : -1;
	if (s0 > 1e-6) {
		const normal: Vec3 = horizontal ? [-dir, 0, 0] : [0, -dir, 0];
		faces.push({ pts: baseRing, normal, kind: faceKind(normal) });
	}
	if (s1 > 1e-6) {
		const normal: Vec3 = horizontal ? [dir, 0, 0] : [0, dir, 0];
		faces.push({ pts: topRing, normal, kind: faceKind(normal) });
	}
	for (let i = 0; i < ring.length; i++) {
		const next = (i + 1) % ring.length;
		let pts: Vec3[];
		if (s1 <= 1e-6) pts = [baseRing[i]!, baseRing[next]!, apexTop];
		else if (s0 <= 1e-6) pts = [apexBase, topRing[next]!, topRing[i]!];
		else pts = [baseRing[i]!, baseRing[next]!, topRing[next]!, topRing[i]!];
		const normal = outwardNormal(pts, center);
		faces.push({
			pts,
			normal,
			kind: faceKind(normal),
			...(curved ? { fill: curvedShade(meta.color, normal) } : {}),
		});
	}
	return { ...meta, center, faces };
}
