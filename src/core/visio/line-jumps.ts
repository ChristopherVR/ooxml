import type { VisioMatrix, VisioPage, VisioShape } from './model';
import { visioPageLayout } from './page-layout';

/** Default jump width in inches at 1:1 scale, measured in Visio 16 (factor * routing spacing). */
export const VISIO_LINE_JUMP_WIDTH = (2 / 3) * 0.125;
/** Connectors and crossings beyond these are drawn without jumps. */
export const VISIO_LINE_JUMP_LIMITS = Object.freeze({ connectors: 2000, segments: 20_000 });

interface Point {
	x: number;
	y: number;
}
interface Segment {
	a: Point;
	b: Point;
	owner: number;
}
type Polyline = Point[];

const TOKEN = /([MLZmlz]|[A-Za-z])|([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)/g;

/** Polylines of a path made only of absolute M and L commands; undefined for anything else. */
function polylines(path: string): Polyline[] | undefined {
	const result: Polyline[] = [];
	let command = '';
	let pending: number | undefined;
	for (const match of path.matchAll(TOKEN)) {
		if (match[1]) {
			if (match[1] !== 'M' && match[1] !== 'L') return undefined;
			if (pending !== undefined) return undefined;
			command = match[1];
			continue;
		}
		const value = Number(match[2]);
		if (!Number.isFinite(value) || !command) return undefined;
		if (pending === undefined) {
			pending = value;
			continue;
		}
		const point = { x: pending, y: value };
		pending = undefined;
		if (command === 'M') {
			result.push([point]);
			command = 'L';
		} else result.at(-1)?.push(point);
	}
	return pending === undefined && result.every((line) => line.length > 1) ? result : undefined;
}

const apply = (m: VisioMatrix, p: Point): Point => ({
	x: m[0] * p.x + m[2] * p.y + m[4],
	y: m[1] * p.x + m[3] * p.y + m[5],
});
function invert(m: VisioMatrix): VisioMatrix | undefined {
	const det = m[0] * m[3] - m[1] * m[2];
	if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return undefined;
	return [
		m[3] / det,
		-m[1] / det,
		-m[2] / det,
		m[0] / det,
		(m[2] * m[5] - m[3] * m[4]) / det,
		(m[1] * m[4] - m[0] * m[5]) / det,
	];
}
const clean = (value: number) => (Math.abs(value) < 1e-9 ? 0 : Number(value.toFixed(6)));
const text = (p: Point) => `${clean(p.x)} ${clean(p.y)}`;
// Native Visio classifies slanted straight segments by their predominant axis.
const horizontal = (s: Segment) => Math.abs(s.b.x - s.a.x) > Math.abs(s.b.y - s.a.y);
const vertical = (s: Segment) => Math.abs(s.b.y - s.a.y) > Math.abs(s.b.x - s.a.x);

/** Parameter along `s` where `other` properly crosses it, away from both segments' ends. */
function crossing(s: Segment, other: Segment, margin: number): number | undefined {
	const dx = s.b.x - s.a.x,
		dy = s.b.y - s.a.y,
		ox = other.b.x - other.a.x,
		oy = other.b.y - other.a.y;
	const det = dx * oy - dy * ox;
	const length = Math.hypot(dx, dy),
		otherLength = Math.hypot(ox, oy);
	// Parallel or nearly parallel lines do not jump each other.
	if (Math.abs(det) < 1e-9 || Math.abs(det) / (length * otherLength) < 0.2) return undefined;
	const t = ((other.a.x - s.a.x) * oy - (other.a.y - s.a.y) * ox) / det;
	const u = ((other.a.x - s.a.x) * dy - (other.a.y - s.a.y) * dx) / det;
	// A line that only touches the other (a T junction or a shared end) does not jump it.
	if (u * otherLength < 1e-6 || (1 - u) * otherLength < 1e-6) return undefined;
	if (t * length < margin || (1 - t) * length < margin) return undefined;
	return t;
}

/** Which of two crossing segments jumps, by the page's LineJumpCode. */
function jumps(code: number, s: Segment, other: Segment, shapes: readonly VisioShape[]): boolean {
	const own = shapes[s.owner]?.lineJumpCode ?? 0;
	const peer = shapes[other.owner]?.lineJumpCode ?? 0;
	// Measured in native Visio: conflicting Always/Other policies use the page rule.
	if (own === 4 || peer === 4 || own === 1) return false;
	if (own !== peer) {
		if (own === 2 || peer === 3) return true;
		if (own === 3 || peer === 2) return false;
	}
	if (code === 1) return horizontal(s) && !horizontal(other);
	if (code === 2) return vertical(s) && !vertical(other);
	// Routing order is not stored: "last routed" is taken as the later connector in drawing order.
	if (code === 3 || code === 4) return s.owner > other.owner;
	return code === 5 && s.owner < other.owner;
}

/** The drawing commands for one jump from `from` to `to` (page space), mapped to local space. */
function jump(
	style: number,
	from: Point,
	to: Point,
	local: (p: Point) => Point,
	reverse: boolean,
): string {
	const half = Math.hypot(to.x - from.x, to.y - from.y) / 2;
	const ux = (to.x - from.x) / (2 * half),
		uy = (to.y - from.y) / (2 * half);
	// Bulge up for a horizontal line and to the left for a vertical one.
	let nx = -uy,
		ny = ux;
	if (ny < -1e-9 || (Math.abs(ny) <= 1e-9 && nx > 0)) {
		nx = -nx;
		ny = -ny;
	}
	if (reverse) {
		nx = -nx;
		ny = -ny;
	}
	const centre = { x: (from.x + to.x) / 2, y: (from.y + to.y) / 2 };
	const at = (along: number, out: number): Point =>
		local({ x: centre.x + ux * along + nx * out, y: centre.y + uy * along + ny * out });
	const end = local(to);
	if (style === 1) return `M ${text(end)}`;
	if (style === 2) return `L ${text(at(-half, half))} L ${text(at(half, half))} L ${text(end)}`;
	if (style >= 3) {
		// Styles 3 to 8 are two to seven straight sides over the same half circle.
		const sides = style - 1;
		let out = '';
		for (let index = 1; index < sides; ++index) {
			const angle = Math.PI - (Math.PI * index) / sides;
			out += `L ${text(at(Math.cos(angle) * half, Math.sin(angle) * half))} `;
		}
		return `${out}L ${text(end)}`;
	}
	const start = local(from),
		top = at(0, half);
	const radius = Math.hypot(end.x - start.x, end.y - start.y) / 2;
	const sweep =
		(end.x - start.x) * (top.y - start.y) - (end.y - start.y) * (top.x - start.x) > 0 ? 0 : 1;
	return `A ${clean(radius)} ${clean(radius)} 0 0 ${sweep} ${text(end)}`;
}

/** Routable connectors jump; a line drawn with the Line tool neither jumps nor is jumped. */
const connector = (shape: VisioShape, glued: ReadonlySet<string>) =>
	shape.kind === 'connector' &&
	!shape.hidden &&
	!shape.children.length &&
	(shape.connectorRoute !== undefined || glued.has(shape.id));

/**
 * Line jumps of a page: for each top-level connector that crosses another, its geometry paths
 * with a jump at every crossing, by the page's LineJumpCode and LineJumpStyle. Only routable or
 * glued connectors drawn with straight segments take part; curved ones neither jump nor are jumped. Paths are in
 * the shape's own coordinates, index for index with `shape.geometry`; a shape without crossings
 * is not in the map.
 */
export function visioLineJumpPaths(page: VisioPage): Map<VisioShape, string[]> {
	const result = new Map<VisioShape, string[]>();
	const layout = visioPageLayout(page);
	const code = layout.lineJumpCode;
	const glued = new Set(page.connectors.map((connection) => connection.fromShapeId));
	const shapes = page.shapes
		.filter((shape) => connector(shape, glued))
		.slice(0, VISIO_LINE_JUMP_LIMITS.connectors);
	// Visio persists jump decorations inside a dynamic connector's Geometry rows.
	// A straight route's endpoints are authoritative: rebuild its baseline before
	// applying current jump settings, including when jumps have been turned off.
	for (const shape of shapes) {
		if (shape.connectorRoute !== 'straight' || !shape.lineEnds) continue;
		if (shape.geometry.length !== 1) continue;
		const geometry = shape.geometry[0]!;
		if (!geometry.stroke || geometry.fill) continue;
		const { begin, end } = shape.lineEnds;
		if (![begin.x, begin.y, end.x, end.y].every(Number.isFinite)) continue;
		const path = `M ${text(begin)} L ${text(end)}`;
		if (geometry.path !== path) result.set(shape, [path]);
	}
	if (shapes.length < 2) return result;
	if (
		(code < 1 || code > 5) &&
		!shapes.some((shape) => shape.lineJumpCode === 2 || shape.lineJumpCode === 3)
	)
		return result;
	const lines: { shape: VisioShape; owner: number; geometry: number; points: Polyline }[] = [];
	const segments: Segment[] = [];
	shapes.forEach((shape, owner) => {
		shape.geometry.forEach((geometry, index) => {
			if (!geometry.stroke || geometry.fill) return;
			for (const line of polylines(result.get(shape)?.[index] ?? geometry.path) ?? []) {
				const points = line.map((point) => apply(shape.transform, point));
				lines.push({ shape, owner, geometry: index, points });
				for (let i = 1; i < points.length; ++i)
					segments.push({ a: points[i - 1]!, b: points[i]!, owner });
			}
		});
	});
	if (segments.length > VISIO_LINE_JUMP_LIMITS.segments) return result;
	const bounded = (value: number | undefined, fallback: number, max: number) =>
		typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= max
			? value
			: fallback;
	const scale = bounded(page.drawingToPageScale, 1, Number.MAX_VALUE);
	const halfX =
		(bounded(layout.lineJumpFactorX, 2 / 3, 10) * bounded(layout.lineToLineX, 0.125, 1e4) * scale) /
		2;
	const halfY =
		(bounded(layout.lineJumpFactorY, 2 / 3, 10) * bounded(layout.lineToLineY, 0.125, 1e4) * scale) /
		2;
	const paths = new Map<VisioShape, Map<number, string[]>>();
	const any = new Set<VisioShape>();
	for (const line of lines) {
		const inverse = invert(line.shape.transform);
		if (!inverse) continue;
		const local = (point: Point) => apply(inverse, point);
		let out = `M ${text(local(line.points[0]!))}`;
		for (let i = 1; i < line.points.length; ++i) {
			const s: Segment = { a: line.points[i - 1]!, b: line.points[i]!, owner: line.owner };
			const length = Math.hypot(s.b.x - s.a.x, s.b.y - s.a.y);
			const half = vertical(s) ? halfY : halfX;
			if (half <= 0) {
				out += ` L ${text(local(s.b))}`;
				continue;
			}
			const hits = segments
				.filter((other) => other.owner !== s.owner && jumps(code, s, other, shapes))
				.flatMap((other) => crossing(s, other, half * 1.5) ?? [])
				.sort((a, b) => a - b);
			let last = -Infinity;
			for (const t of hits) {
				// Crossings closer than one jump share the first jump.
				if ((t - last) * length < half * 2.5) continue;
				last = t;
				const at = (offset: number): Point => ({
					x: s.a.x + ((s.b.x - s.a.x) * (t * length + offset)) / length,
					y: s.a.y + ((s.b.y - s.a.y) * (t * length + offset)) / length,
				});
				const override = line.shape.lineJumpStyle ?? 0;
				const savedStyle = override >= 1 && override <= 9 ? override : layout.lineJumpStyle;
				const style = savedStyle >= 1 && savedStyle <= 9 ? savedStyle - 1 : 0;
				const direction = horizontal(s)
					? line.shape.lineJumpDirectionX || layout.lineJumpDirectionX
					: vertical(s)
						? line.shape.lineJumpDirectionY || layout.lineJumpDirectionY
						: 0;
				out += ` L ${text(local(at(-half)))} ${jump(style, at(-half), at(half), local, direction === 2)}`;
				any.add(line.shape);
			}
			out += ` L ${text(local(s.b))}`;
		}
		const byGeometry = paths.get(line.shape) ?? new Map<number, string[]>();
		byGeometry.set(line.geometry, [...(byGeometry.get(line.geometry) ?? []), out]);
		paths.set(line.shape, byGeometry);
	}
	for (const shape of any) {
		const byGeometry = paths.get(shape)!;
		result.set(
			shape,
			shape.geometry.map(
				(geometry, index) =>
					byGeometry.get(index)?.join(' ') ?? result.get(shape)?.[index] ?? geometry.path,
			),
		);
	}
	return result;
}
