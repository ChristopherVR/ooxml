/**
 * Perspective projection and extrusion of a flat shape turned by Visio's 3-D Rotation, DOM-free.
 *
 * The rotation order and signs were fitted to Visio 16's raster export: a point (x, y, z) in the
 * shape's y-up frame turns about the vertical axis by RotationXAngle (positive brings the right
 * side towards the viewer), then about the horizontal axis by RotationYAngle (positive brings the
 * top towards the viewer), then in the picture plane by RotationZAngle (counter-clockwise).
 * `Perspective` is the camera's field of view: a point `z` nearer the viewer is scaled by
 * f / (f - z) with f = 6.36 in / tan(fov / 2), a distance independent of the shape's size that was
 * measured at 30, 60 and 100 degrees (Visio clamps the field of view at 120 degrees). Extrusion
 * (`BevelDepthSize`) runs away from the viewer behind the shape, so its side faces are the quads
 * between the front outline and the back outline.
 * Lighting, materials and Visio's shading of the sides are not reproduced.
 */
import { parsePath } from './path-segments';

export interface Point2 {
	x: number;
	y: number;
}
export interface Turn3d {
	x: number;
	y: number;
	z: number;
	perspective: number;
}
export const VISIO_PERSPECTIVE_LIMIT = (120 * Math.PI) / 180;
/** Camera distance in inches at a field of view of 90 degrees. */
export const VISIO_CAMERA_INCHES = 6.36;
const CURVE_STEPS = 12;

export interface Projection {
	/** The picture position of the shape point (x, y) lifted by `z` towards the viewer. */
	point(x: number, y: number, z?: number): Point2;
	/** Distance towards the viewer of the same point; larger is nearer. */
	depth(x: number, y: number, z?: number): number;
	/** The picture-plane direction and viewer-facing part of a local direction (x, y, 0). */
	direction(x: number, y: number): { x: number; y: number; toward: number };
}

/** `unit` is how many inches one local unit measures; `centre` is the turning point. */
export function createProjection(turn: Turn3d, centre: Point2, unit = 1): Projection {
	const [ca, sa] = [Math.cos(turn.x), Math.sin(turn.x)];
	const [cb, sb] = [Math.cos(turn.y), Math.sin(turn.y)];
	const [cc, sc] = [Math.cos(turn.z), Math.sin(turn.z)];
	const fov = Math.min(Math.max(turn.perspective, 0), VISIO_PERSPECTIVE_LIMIT);
	const camera = fov > 0.001 ? VISIO_CAMERA_INCHES / Math.tan(fov / 2) / (unit || 1) : undefined;
	const turned = (x: number, y: number, z: number) => {
		const x1 = x * ca - z * sa;
		const z1 = x * sa + z * ca;
		return { x: x1, y: y * cb - z1 * sb, z: y * sb + z1 * cb };
	};
	return {
		point(x, y, z = 0) {
			const t = turned(x - centre.x, y - centre.y, z);
			const scale = camera === undefined ? 1 : camera / Math.max(camera - t.z, camera * 0.05);
			const [px, py] = [t.x * scale, t.y * scale];
			return { x: centre.x + cc * px - sc * py, y: centre.y + sc * px + cc * py };
		},
		depth: (x, y, z = 0) => turned(x - centre.x, y - centre.y, z).z,
		direction(x, y) {
			const t = turned(x, y, 0);
			return { x: cc * t.x - sc * t.y, y: sc * t.x + cc * t.y, toward: t.z };
		},
	};
}

export interface FlatSubpath {
	points: Point2[];
	closed: boolean;
}
/** Path data flattened to polylines (curves and arcs in 12 steps each); undefined if malformed. */
export function flattenPath(d: string): FlatSubpath[] | undefined {
	const segments = parsePath(d);
	if (!segments) return undefined;
	const result: FlatSubpath[] = [];
	let current: FlatSubpath | undefined;
	for (const segment of segments) {
		if (segment.t === 'M') {
			current = { points: [{ x: segment.x, y: segment.y }], closed: false };
			result.push(current);
		} else if (!current) return undefined;
		else if (segment.t === 'L') current.points.push({ x: segment.x, y: segment.y });
		else if (segment.t === 'Z') {
			current.closed = true;
			const first = current.points[0]!;
			current = { points: [first], closed: false };
			// A segment after Z continues from the start without a new M.
			result.push(current);
		} else {
			const from = current.points[current.points.length - 1]!;
			for (let step = 1; step <= CURVE_STEPS; step++) {
				const t = step / CURVE_STEPS;
				const u = 1 - t;
				const [a, b, c, e] = [u * u * u, 3 * u * u * t, 3 * u * t * t, t * t * t];
				current.points.push({
					x: a * from.x + b * segment.x1 + c * segment.x2 + e * segment.x,
					y: a * from.y + b * segment.y1 + c * segment.y2 + e * segment.y,
				});
			}
		}
	}
	// Visio geometry often closes a filled outline by returning to its start without a Z.
	for (const sub of result) {
		const [first, last] = [sub.points[0]!, sub.points[sub.points.length - 1]!];
		if (
			sub.points.length > 3 &&
			Math.abs(first.x - last.x) < 1e-9 &&
			Math.abs(first.y - last.y) < 1e-9
		)
			sub.closed = true;
	}
	return result.filter((sub) => sub.closed || sub.points.length > 1);
}

const fixed = (value: number) => String(Math.round(value * 1e5) / 1e5);
const polygonData = (points: readonly Point2[], closed: boolean) =>
	`${points.map((p, i) => `${i ? 'L' : 'M'} ${fixed(p.x)} ${fixed(p.y)}`).join(' ')}${closed ? ' Z' : ''}`;
/** The path projected point by point, as polylines; curves are flattened. */
export function projectPathData(d: string, projection: Projection): string | undefined {
	const subpaths = flattenPath(d);
	if (!subpaths) return undefined;
	return subpaths
		.map((sub) =>
			polygonData(
				sub.points.map((p) => projection.point(p.x, p.y)),
				sub.closed,
			),
		)
		.join(' ');
}

export interface ExtrusionFace {
	points: Point2[];
	/** Brightness factor for the face colour: below 1 darkens, above 1 lightens. */
	shade: number;
	depth: number;
}
export interface Extrusion {
	/** Back outline (shape behind the front), one polygon per closed subpath. */
	back: Point2[][];
	/** Visible side faces, farthest first. */
	sides: ExtrusionFace[];
}
const area = (points: readonly Point2[]) => {
	let sum = 0;
	for (let i = 0; i < points.length; i++) {
		const [p, q] = [points[i]!, points[(i + 1) % points.length]!];
		sum += p.x * q.y - q.x * p.y;
	}
	return sum / 2;
};
/**
 * Side faces of the shape pushed `depth` local units away from the viewer. Brightness is fitted to
 * four of Visio's renders (right side dark, top slightly dark, bottom light, left near the colour)
 * and is not Visio's lighting model.
 */
export function extrudePaths(
	paths: readonly string[],
	projection: Projection,
	depth: number,
): Extrusion | undefined {
	const back: Point2[][] = [];
	const sides: ExtrusionFace[] = [];
	for (const d of paths) {
		const subpaths = flattenPath(d);
		if (!subpaths) return undefined;
		for (const sub of subpaths) {
			const points = [...sub.points];
			if (!sub.closed || points.length < 3) continue;
			// Drop the repeated closing point so no zero-length edge is produced.
			const first = points[0]!;
			const last = points[points.length - 1]!;
			if (Math.abs(first.x - last.x) < 1e-9 && Math.abs(first.y - last.y) < 1e-9) points.pop();
			const turnSign = area(points) >= 0 ? 1 : -1;
			back.push(points.map((p) => projection.point(p.x, p.y, -depth)));
			for (let i = 0; i < points.length; i++) {
				const [p, q] = [points[i]!, points[(i + 1) % points.length]!];
				const [dx, dy] = [q.x - p.x, q.y - p.y];
				const length = Math.hypot(dx, dy);
				if (length < 1e-9) continue;
				const normal = projection.direction((turnSign * dy) / length, (-turnSign * dx) / length);
				if (normal.toward <= 1e-6) continue;
				const screen = Math.hypot(normal.x, normal.y) || 1;
				const shade = 0.9 - 0.18 * (normal.y / screen) - 0.7 * Math.max(0, normal.x / screen);
				sides.push({
					points: [
						projection.point(p.x, p.y),
						projection.point(q.x, q.y),
						projection.point(q.x, q.y, -depth),
						projection.point(p.x, p.y, -depth),
					],
					shade: Math.max(0.1, shade),
					depth: projection.depth((p.x + q.x) / 2, (p.y + q.y) / 2, -depth / 2),
				});
			}
		}
	}
	sides.sort((a, b) => a.depth - b.depth);
	return { back, sides };
}
export const extrusionPathData = (polygons: readonly (readonly Point2[])[]): string =>
	polygons.map((points) => polygonData(points, true)).join(' ');
