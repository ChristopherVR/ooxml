// DOM-free 3-D meshes of the chart marks. The scene frame is x right, y up, z away from the
// viewer. Every mark is a prism: a polygon in the xy plane extruded along z (a bar is a rectangle,
// a line ribbon a parallelogram per segment, an area its outline). Faces carry their outward
// normal so the painter can cull the ones facing away.
import type { Chart3DFace } from './chart-3d-shading';

export type Vec3 = readonly [number, number, number];

export interface Chart3DMeshFace {
	pts: Vec3[];
	/** Outward unit normal; undefined for a double-sided face (surface patches). */
	normal: Vec3 | undefined;
	kind: Chart3DFace;
	/** A colour replacing the mesh colour (surface bands shade per face). */
	fill?: string;
}

export interface Chart3DMesh {
	series: number;
	point: number;
	color: string;
	center: Vec3;
	faces: Chart3DMeshFace[];
}

export interface Chart3DMeshMeta {
	series: number;
	point: number;
	color: string;
}

/** The shading orientation of a face from its dominant normal axis. */
export function faceKind(n: Vec3): Chart3DFace {
	const [x, y, z] = n;
	if (Math.abs(y) >= Math.abs(x) && Math.abs(y) >= Math.abs(z)) return y >= 0 ? 'top' : 'bottom';
	if (Math.abs(x) >= Math.abs(z)) return x >= 0 ? 'right' : 'left';
	return z < 0 ? 'front' : 'back';
}

function signedArea(outline: readonly (readonly [number, number])[]): number {
	let sum = 0;
	for (let i = 0; i < outline.length; i++) {
		const [x0, y0] = outline[i]!;
		const [x1, y1] = outline[(i + 1) % outline.length]!;
		sum += x0 * y1 - x1 * y0;
	}
	return sum / 2;
}

/**
 * A polygon in the xy plane extruded from `z0` to `z1`. Returns undefined for a degenerate outline
 * (no area), which a zero-valued bar or a flat area produces.
 */
export function extrudeMesh(
	outline: readonly (readonly [number, number])[],
	z0: number,
	z1: number,
	meta: Chart3DMeshMeta,
): Chart3DMesh | undefined {
	if (outline.length < 3 || Math.abs(signedArea(outline)) < 1e-6) return undefined;
	const ccw = signedArea(outline) > 0 ? [...outline] : [...outline].reverse();
	const faces: Chart3DMeshFace[] = [
		{ pts: ccw.map(([x, y]) => [x, y, z0] as Vec3), normal: [0, 0, -1], kind: 'front' },
		{ pts: ccw.map(([x, y]) => [x, y, z1] as Vec3), normal: [0, 0, 1], kind: 'back' },
	];
	for (let i = 0; i < ccw.length; i++) {
		const [ax, ay] = ccw[i]!;
		const [bx, by] = ccw[(i + 1) % ccw.length]!;
		const dx = bx - ax;
		const dy = by - ay;
		const len = Math.hypot(dx, dy);
		if (len < 1e-9) continue;
		const normal: Vec3 = [dy / len, -dx / len, 0];
		faces.push({
			pts: [
				[ax, ay, z0],
				[bx, by, z0],
				[bx, by, z1],
				[ax, ay, z1],
			],
			normal,
			kind: faceKind(normal),
		});
	}
	const xs = ccw.map((p) => p[0]);
	const ys = ccw.map((p) => p[1]);
	const center: Vec3 = [
		(Math.min(...xs) + Math.max(...xs)) / 2,
		(Math.min(...ys) + Math.max(...ys)) / 2,
		(z0 + z1) / 2,
	];
	return { ...meta, center, faces };
}

/** An axis-aligned box. */
export function boxMesh(
	x0: number,
	x1: number,
	y0: number,
	y1: number,
	z0: number,
	z1: number,
	meta: Chart3DMeshMeta,
): Chart3DMesh | undefined {
	return extrudeMesh(
		[
			[Math.min(x0, x1), Math.min(y0, y1)],
			[Math.max(x0, x1), Math.min(y0, y1)],
			[Math.max(x0, x1), Math.max(y0, y1)],
			[Math.min(x0, x1), Math.max(y0, y1)],
		],
		Math.min(z0, z1),
		Math.max(z0, z1),
		meta,
	);
}
