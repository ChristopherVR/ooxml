// The camera of a 3-D chart: scene points (x right, y up, z away) to SVG pixels, the view
// direction used to cull back faces, and the far-to-near order of marks. The 3-D box is fitted
// uniformly into the plot rectangle, so the same scene works for every frame size.
import type { Rect } from './chart-svg-util';
import type { Chart3DProjection } from './chart-3d-projection';
import type { Vec3 } from './chart-3d-mesh';

export interface Chart3DSize {
	w: number;
	h: number;
	d: number;
}

export interface Chart3DCamera {
	/** Scene point to SVG pixels; `depth` grows with the distance from the viewer. */
	project(p: Vec3): { x: number; y: number; depth: number };
	/** The direction the viewer looks along at a scene point (a face is visible when n.d < 0). */
	viewDir(p: Vec3): Vec3;
	/** The eye position of a perspective camera. */
	eye: Vec3 | undefined;
	/** Negative when `a` is drawn before (is farther than) `b`. */
	order(a: Vec3, b: Vec3): number;
	/** Whether the x = 0 wall faces the viewer from inside (the camera is right of centre). */
	sideWallAtZero: boolean;
}

/** Shear applied per unit of depth by the oblique projection (a cavalier projection halves it). */
const OBLIQUE_DEPTH_SCALE = 0.5;
const rad = (deg: number): number => (deg * Math.PI) / 180;

/** Fit the projected extremes of the box into `fit`, centred, keeping the aspect ratio. */
function fitTransform(
	raw: (p: Vec3) => { x: number; y: number },
	size: Chart3DSize,
	fit: Rect,
): { scale: number; ox: number; oy: number } {
	const xs: number[] = [];
	const ys: number[] = [];
	for (const x of [0, size.w])
		for (const y of [0, size.h])
			for (const z of [0, size.d]) {
				const q = raw([x, y, z]);
				xs.push(q.x);
				ys.push(q.y);
			}
	const minX = Math.min(...xs);
	const maxX = Math.max(...xs);
	const minY = Math.min(...ys);
	const maxY = Math.max(...ys);
	const scale = Math.min(fit.w / Math.max(1e-6, maxX - minX), fit.h / Math.max(1e-6, maxY - minY));
	return {
		scale,
		ox: fit.x + (fit.w - (maxX - minX) * scale) / 2 - minX * scale,
		oy: fit.y + (fit.h - (maxY - minY) * scale) / 2 + maxY * scale,
	};
}

export function createChart3DCamera(
	projection: Chart3DProjection,
	size: Chart3DSize,
	fit: Rect,
): Chart3DCamera {
	if (projection.mode === 'oblique') {
		const kx = projection.shearX * OBLIQUE_DEPTH_SCALE;
		const ky = projection.shearY * OBLIQUE_DEPTH_SCALE;
		const raw = (p: Vec3) => ({ x: p[0] + p[2] * kx, y: p[1] - p[2] * ky });
		const t = fitTransform(raw, size, fit);
		return {
			project: (p) => {
				const q = raw(p);
				return { x: t.ox + q.x * t.scale, y: t.oy - q.y * t.scale, depth: p[2] };
			},
			viewDir: () => [-kx, ky, 1],
			eye: undefined,
			order: (a, b) =>
				b[2] - a[2] ||
				(projection.shearX >= 0 ? a[0] - b[0] : b[0] - a[0]) ||
				(projection.shearY <= 0 ? a[1] - b[1] : b[1] - a[1]),
			sideWallAtZero: projection.shearX >= 0,
		};
	}
	const cx = size.w / 2;
	const cy = size.h / 2;
	const cz = size.d / 2;
	const [cosY, sinY] = [Math.cos(rad(projection.rotYDeg)), Math.sin(rad(projection.rotYDeg))];
	const [cosX, sinX] = [Math.cos(rad(projection.rotXDeg)), Math.sin(rad(projection.rotXDeg))];
	const radius = Math.hypot(size.w, size.h, size.d) / 2;
	const dist = radius / Math.tan(rad(projection.fovDeg) / 2);
	const rotate = (p: Vec3): Vec3 => {
		const x = p[0] - cx;
		const y = p[1] - cy;
		const z = p[2] - cz;
		const x1 = x * cosY + z * sinY;
		const z1 = -x * sinY + z * cosY;
		return [x1, y * cosX + z1 * sinX, -y * sinX + z1 * cosX];
	};
	const raw = (p: Vec3) => {
		const r = rotate(p);
		const s = dist / (dist + r[2]);
		return { x: r[0] * s, y: r[1] * s, depth: r[2] };
	};
	const t = fitTransform(raw, size, fit);
	// The eye sits at (0, 0, -dist) in the rotated frame: undo the elevation, then the azimuth.
	const ey = dist * sinX;
	const ez1 = -dist * cosX;
	const eyeX = cx - ez1 * sinY;
	const eye: Vec3 = [eyeX, cy + ey, cz + ez1 * cosY];
	return {
		project: (p) => {
			const q = raw(p);
			return { x: t.ox + q.x * t.scale, y: t.oy - q.y * t.scale, depth: q.depth };
		},
		viewDir: (p) => [p[0] - eye[0], p[1] - eye[1], p[2] - eye[2]],
		eye,
		order: (a, b) => {
			const da = (a[0] - eyeX) ** 2 + (a[1] - eye[1]) ** 2 + (a[2] - eye[2]) ** 2;
			const db = (b[0] - eyeX) ** 2 + (b[1] - eye[1]) ** 2 + (b[2] - eye[2]) ** 2;
			return db - da;
		},
		sideWallAtZero: eyeX > cx,
	};
}
