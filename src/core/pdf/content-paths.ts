/**
 * Path construction and painting (ISO 32000-1 section 8.5): collects a path's subpaths and, when
 * it is painted, sorts it into straight rules and rectangles (`PdfPageRule`, which table, underline
 * and shading detection read) or a whole path with its geometry (`PdfPagePath`) for curves,
 * diagonals and polygons. Paths painted with patterns or cut by a clip are not returned as paths.
 */
import { apply, multiply, type Matrix } from './matrix';
import type {
	PdfGradientFill,
	PdfPagePath,
	PdfPageRule,
	PdfPathSegment,
	PdfPatternFill,
} from './page-content';

/** The clipping region's bounds in reading coordinates, and whether it is exactly that box. */
export interface ClipBox {
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	rect: boolean;
}

export interface PaintStyle {
	ctm: Matrix;
	fill: string;
	stroke: string;
	lineWidth: number;
	fillAlpha: number;
	strokeAlpha: number;
	dash: number[];
	fillPattern: boolean;
	strokePattern: boolean;
	fillGradient?: PdfGradientFill | undefined;
	fillPatternFill?: PdfPatternFill | undefined;
	clip: ClipBox | undefined;
}

export type PaintResult =
	| { rules: PdfPageRule[] }
	| { path: PdfPagePath }
	| { unsupported: 'pattern' | 'clip' | 'size' };

type Point = [number, number];
interface Subpath {
	segments: PdfPathSegment[];
	points: Point[];
	curves: boolean;
	closed: boolean;
	rect: boolean;
}

/** Segments one path may hold before it is too large to rebuild. */
const MAX_SEGMENTS = 4000;
/** How far (points) a path may reach past a clip before it counts as cut. */
const CLIP_TOLERANCE = 1;

const near = (a: number, b: number, tolerance = 0.01) => Math.abs(a - b) <= tolerance;

export class PathBuilder {
	#subpaths: Subpath[] = [];
	#open: Subpath | undefined;
	#count = 0;
	/** The current point in user space. */
	current: Point | undefined;

	#start(x: number, y: number): Subpath {
		const sub: Subpath = {
			segments: [],
			points: [[x, y]],
			curves: false,
			closed: false,
			rect: false,
		};
		sub.segments.push({ op: 'M', x, y });
		this.#subpaths.push(sub);
		this.#open = sub;
		this.#count++;
		return sub;
	}

	#subpath(): Subpath | undefined {
		if (this.#open && !this.#open.closed) return this.#open;
		if (!this.current) return undefined;
		return this.#start(this.current[0], this.current[1]);
	}

	moveTo(x: number, y: number): void {
		this.#open = undefined;
		this.current = [x, y];
		this.#start(x, y);
	}

	lineTo(x: number, y: number): void {
		const sub = this.#subpath();
		if (sub) {
			sub.segments.push({ op: 'L', x, y });
			sub.points.push([x, y]);
			this.#count++;
		}
		this.current = [x, y];
	}

	curveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): void {
		const sub = this.#subpath();
		if (sub) {
			sub.segments.push({ op: 'C', x1, y1, x2, y2, x, y });
			sub.points.push([x1, y1], [x2, y2], [x, y]);
			sub.curves = true;
			this.#count++;
		}
		this.current = [x, y];
	}

	rect(x: number, y: number, width: number, height: number): void {
		const sub = this.#start(x, y);
		for (const [px, py] of [
			[x + width, y],
			[x + width, y + height],
			[x, y + height],
		] as Point[]) {
			sub.segments.push({ op: 'L', x: px, y: py });
			sub.points.push([px, py]);
		}
		sub.segments.push({ op: 'Z' });
		sub.closed = true;
		sub.rect = true;
		this.#count += 4;
		this.current = [x, y];
	}

	close(): void {
		const sub = this.#open;
		if (!sub || sub.closed) return;
		sub.segments.push({ op: 'Z' });
		sub.closed = true;
		this.#count++;
		this.current = sub.points[0];
	}

	clear(): void {
		this.#subpaths = [];
		this.#open = undefined;
		this.#count = 0;
		this.current = undefined;
	}

	/** The bounds of the path in reading coordinates, for `W`, without painting it. */
	bounds(page: Matrix, ctm: Matrix): ClipBox | undefined {
		const m = multiply(ctm, page);
		const points = this.#subpaths.flatMap((sub) => sub.points.map(([x, y]) => apply(m, x, y)));
		if (!points.length) return undefined;
		const xs = points.map((point) => point[0]);
		const ys = points.map((point) => point[1]);
		const rect =
			this.#subpaths.length === 1 &&
			!this.#subpaths[0]!.curves &&
			rectangle(this.#subpaths[0]!.points.map(([x, y]) => apply(m, x, y)));
		return {
			x0: Math.min(...xs),
			y0: Math.min(...ys),
			x1: Math.max(...xs),
			y1: Math.max(...ys),
			rect,
		};
	}

	/** Sorts the painted path into rules or a whole path, then starts a new path. */
	paint(
		page: Matrix,
		style: PaintStyle,
		stroked: boolean,
		filled: boolean,
		evenOdd: boolean,
	): PaintResult | undefined {
		const subpaths = this.#subpaths;
		const count = this.#count;
		this.clear();
		if (!subpaths.length) return undefined;
		const m = multiply(style.ctm, page);
		const scale = Math.hypot(m[0], m[1]) || 1;
		const lineWidth = style.lineWidth * scale;
		const reading = subpaths.map((sub) => sub.points.map(([x, y]) => apply(m, x, y)));
		const all = reading.flat();
		const box = {
			x0: Math.min(...all.map((point) => point[0])),
			y0: Math.min(...all.map((point) => point[1])),
			x1: Math.max(...all.map((point) => point[0])),
			y1: Math.max(...all.map((point) => point[1])),
		};
		const clip = style.clip;
		const reach = stroked ? lineWidth / 2 : 0;
		const outsideRectClip =
			!!clip?.rect &&
			(box.x1 + reach <= clip.x0 ||
				box.x0 - reach >= clip.x1 ||
				box.y1 + reach <= clip.y0 ||
				box.y0 - reach >= clip.y1);
		if (outsideRectClip) return { rules: [] };
		const rectClipped =
			!!clip?.rect &&
			(box.x0 - reach < clip.x0 - CLIP_TOLERANCE ||
				box.y0 - reach < clip.y0 - CLIP_TOLERANCE ||
				box.x1 + reach > clip.x1 + CLIP_TOLERANCE ||
				box.y1 + reach > clip.y1 + CLIP_TOLERANCE);
		const clipped = !!clip && !clip.rect;
		const pattern = (filled && style.fillPattern) || (stroked && style.strokePattern);
		const rules: PdfPageRule[] = [];
		let complex = false;
		const rule = (points: Point[], kind: 'line' | 'rect', fill: boolean) => {
			const xs = points.map((point) => point[0]);
			const ys = points.map((point) => point[1]);
			let x0 = Math.min(...xs);
			let y0 = Math.min(...ys);
			let x1 = Math.max(...xs);
			let y1 = Math.max(...ys);
			if (clip?.rect && rectClipped) {
				x0 = Math.max(x0, clip.x0);
				y0 = Math.max(y0, clip.y0);
				x1 = Math.min(x1, clip.x1);
				y1 = Math.min(y1, clip.y1);
				if (kind === 'rect' ? x1 <= x0 || y1 <= y0 : x1 < x0 || y1 < y0) return;
			}
			const item: PdfPageRule = {
				x0,
				y0,
				x1,
				y1,
				kind,
				stroked,
				filled: fill,
				color: fill ? style.fill : style.stroke,
				lineWidth,
			};
			if (fill && stroked && style.stroke !== style.fill) item.strokeColor = style.stroke;
			const alpha = fill ? style.fillAlpha : style.strokeAlpha;
			if (alpha < 1) item.alpha = alpha;
			if (stroked && style.dash.length) item.dash = style.dash.map((value) => value * scale);
			if (fill && style.fillGradient) item.gradient = style.fillGradient;
			if (fill && style.fillPatternFill) item.patternFill = style.fillPatternFill;
			if (pattern) item.pattern = true;
			if (clipped) item.clipped = true;
			rules.push(item);
		};
		reading.forEach((points, index) => {
			const sub = subpaths[index]!;
			if (complex) return;
			if (sub.curves) complex = true;
			else if (sub.rect || (filled && rectangle(points)))
				if (rectangle(points)) rule(points, 'rect', filled);
				else complex = true;
			else if (filled && area(points) > 0.5) complex = true;
			else {
				const lines = points.slice(1).map((point, at) => [points[at]!, point] as [Point, Point]);
				if (sub.closed && stroked && points.length > 2) lines.push([points.at(-1)!, points[0]!]);
				for (const [a, b] of lines) {
					const dx = Math.abs(a[0] - b[0]);
					const dy = Math.abs(a[1] - b[1]);
					if (dx > Math.max(0.01, dy * 0.005) && dy > Math.max(0.01, dx * 0.005)) complex = true;
				}
				if (!complex) for (const [a, b] of lines) rule([a, b], 'line', filled);
			}
		});
		if (!complex) return { rules };
		if (pattern) return { unsupported: 'pattern' };
		if (clipped) return { unsupported: 'clip' };
		if (count > MAX_SEGMENTS) return { unsupported: 'size' };
		let segments = subpaths.flatMap((sub) =>
			sub.segments.map((segment): PdfPathSegment => {
				if (segment.op === 'Z') return segment;
				const [x, y] = apply(m, segment.x, segment.y);
				if (segment.op !== 'C') return { op: segment.op, x, y };
				const [x1, y1] = apply(m, segment.x1, segment.y1);
				const [x2, y2] = apply(m, segment.x2, segment.y2);
				return { op: 'C', x1, y1, x2, y2, x, y };
			}),
		);
		let outBox = box;
		if (clip?.rect && rectClipped) {
			const allPolygons = subpaths.every((sub) => !sub.curves && sub.closed);
			if (allPolygons && filled) {
				const clippedSegs: PdfPathSegment[] = [];
				for (const points of reading) {
					const poly = clipPolygonToRect(points, clip);
					if (poly.length < 3) continue;
					clippedSegs.push({ op: 'M', x: poly[0]![0], y: poly[0]![1] });
					for (let i = 1; i < poly.length; i++)
						clippedSegs.push({ op: 'L', x: poly[i]![0], y: poly[i]![1] });
					clippedSegs.push({ op: 'Z' });
				}
				if (!clippedSegs.length) return { rules: [] };
				segments = clippedSegs;
			}
			outBox = {
				x0: Math.max(box.x0, clip.x0),
				y0: Math.max(box.y0, clip.y0),
				x1: Math.min(box.x1, clip.x1),
				y1: Math.min(box.y1, clip.y1),
			};
		}
		const path: PdfPagePath = { segments, ...outBox, lineWidth };
		if (filled) path.fill = style.fill;
		if (filled && style.fillGradient) path.gradient = style.fillGradient;
		if (filled && style.fillPatternFill) path.patternFill = style.fillPatternFill;
		if (stroked) path.stroke = style.stroke;
		if (filled && style.fillAlpha < 1) path.fillAlpha = style.fillAlpha;
		if (stroked && style.strokeAlpha < 1) path.strokeAlpha = style.strokeAlpha;
		if (stroked && style.dash.length) path.dash = style.dash.map((value) => value * scale);
		if (filled && evenOdd) path.evenOdd = true;
		return { path };
	}
}

/** Sutherland-Hodgman polygon clipping against an axis-aligned rectangle. */
function clipPolygonToRect(points: Point[], clip: ClipBox): Point[] {
	let poly = [...points];
	if (
		poly.length > 2 &&
		near(poly[0]![0], poly[poly.length - 1]![0]) &&
		near(poly[0]![1], poly[poly.length - 1]![1])
	)
		poly.pop();
	const edges: [(p: Point) => boolean, (a: Point, b: Point) => Point][] = [
		[
			(p) => p[0] >= clip.x0,
			(a, b) => [clip.x0, a[1] + ((clip.x0 - a[0]) / (b[0] - a[0])) * (b[1] - a[1])],
		],
		[
			(p) => p[0] <= clip.x1,
			(a, b) => [clip.x1, a[1] + ((clip.x1 - a[0]) / (b[0] - a[0])) * (b[1] - a[1])],
		],
		[
			(p) => p[1] >= clip.y0,
			(a, b) => [a[0] + ((clip.y0 - a[1]) / (b[1] - a[1])) * (b[0] - a[0]), clip.y0],
		],
		[
			(p) => p[1] <= clip.y1,
			(a, b) => [a[0] + ((clip.y1 - a[1]) / (b[1] - a[1])) * (b[0] - a[0]), clip.y1],
		],
	];
	for (const [inside, intersect] of edges) {
		if (!poly.length) break;
		const next: Point[] = [];
		let prev = poly[poly.length - 1]!;
		let prevIn = inside(prev);
		for (const curr of poly) {
			const currIn = inside(curr);
			if (currIn) {
				if (!prevIn) next.push(intersect(prev, curr));
				next.push(curr);
			} else if (prevIn) next.push(intersect(prev, curr));
			prev = curr;
			prevIn = currIn;
		}
		poly = next;
	}
	return poly;
}

/** Whether the points (a closing repeat allowed) are the corners of an axis-aligned rectangle. */
function rectangle(points: Point[]): boolean {
	const corners = [...points];
	const first = corners[0];
	const last = corners.at(-1);
	if (corners.length === 5 && first && last && near(first[0], last[0]) && near(first[1], last[1]))
		corners.pop();
	if (corners.length !== 4) return false;
	return corners.every((point, index) => {
		const next = corners[(index + 1) % 4]!;
		return near(point[0], next[0], 0.05) || near(point[1], next[1], 0.05);
	});
}

/** The area the polygon through the points encloses (shoelace formula). */
function area(points: Point[]): number {
	let sum = 0;
	points.forEach((point, index) => {
		const next = points[(index + 1) % points.length]!;
		sum += point[0] * next[1] - next[0] * point[1];
	});
	return Math.abs(sum) / 2;
}
