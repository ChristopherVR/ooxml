/**
 * `stroke-clip-outset`: widen a shape's `clip-path` so its centred outline is
 * not cut in half.
 *
 * Every binding clips the shape wrapper to the shape's own geometry (that is
 * how a non-rectangular fill is drawn), and paints the outline as an SVG child
 * of that wrapper. A centred outline straddles the geometry edge, so the clip
 * removed its outer half and a 1pt line on a pin or a chevron drew at 0.5pt.
 *
 * The widened clip is the original path plus one thin quad per outline
 * segment, all wound the same way so the non-zero fill rule takes their union.
 * Curves are flattened first. Fill only reaches the extra band where the
 * outline is solid and covers it, so callers use this for flat, solid lines.
 *
 * @module render/stroke-clip-outset
 */
import { outlinePathData } from './stroke-outline';

type Point = readonly [number, number];

/** Segments a curve is flattened into; the band only needs to hug the edge. */
const CURVE_STEPS = 16;

function flattenArc(
	from: Point,
	rx: number,
	ry: number,
	rotationDeg: number,
	largeArc: boolean,
	sweep: boolean,
	to: Point,
): Point[] {
	if (rx === 0 || ry === 0 || (from[0] === to[0] && from[1] === to[1])) {
		return [to];
	}
	const phi = (rotationDeg * Math.PI) / 180;
	const cosPhi = Math.cos(phi);
	const sinPhi = Math.sin(phi);
	const dx = (from[0] - to[0]) / 2;
	const dy = (from[1] - to[1]) / 2;
	const x1 = cosPhi * dx + sinPhi * dy;
	const y1 = -sinPhi * dx + cosPhi * dy;
	let radiusX = Math.abs(rx);
	let radiusY = Math.abs(ry);
	const lambda = (x1 * x1) / (radiusX * radiusX) + (y1 * y1) / (radiusY * radiusY);
	if (lambda > 1) {
		const scale = Math.sqrt(lambda);
		radiusX *= scale;
		radiusY *= scale;
	}
	const numerator =
		radiusX * radiusX * radiusY * radiusY -
		radiusX * radiusX * y1 * y1 -
		radiusY * radiusY * x1 * x1;
	const denominator = radiusX * radiusX * y1 * y1 + radiusY * radiusY * x1 * x1;
	const coefficient =
		(largeArc === sweep ? -1 : 1) * Math.sqrt(Math.max(0, numerator / denominator));
	const cxp = (coefficient * radiusX * y1) / radiusY;
	const cyp = (-coefficient * radiusY * x1) / radiusX;
	const cx = cosPhi * cxp - sinPhi * cyp + (from[0] + to[0]) / 2;
	const cy = sinPhi * cxp + cosPhi * cyp + (from[1] + to[1]) / 2;
	const angle = (ux: number, uy: number, vx: number, vy: number): number => {
		const sign = ux * vy - uy * vx < 0 ? -1 : 1;
		const dot = (ux * vx + uy * vy) / (Math.hypot(ux, uy) * Math.hypot(vx, vy));
		return sign * Math.acos(Math.max(-1, Math.min(1, dot)));
	};
	const theta1 = angle(1, 0, (x1 - cxp) / radiusX, (y1 - cyp) / radiusY);
	let delta = angle(
		(x1 - cxp) / radiusX,
		(y1 - cyp) / radiusY,
		(-x1 - cxp) / radiusX,
		(-y1 - cyp) / radiusY,
	);
	if (!sweep && delta > 0) {
		delta -= 2 * Math.PI;
	} else if (sweep && delta < 0) {
		delta += 2 * Math.PI;
	}
	const out: Point[] = [];
	for (let step = 1; step <= CURVE_STEPS; step++) {
		const t = theta1 + (delta * step) / CURVE_STEPS;
		const px = radiusX * Math.cos(t);
		const py = radiusY * Math.sin(t);
		out.push([cosPhi * px - sinPhi * py + cx, sinPhi * px + cosPhi * py + cy]);
	}
	return out;
}

/** Flatten absolute `M L H V C Q A Z` path data into polylines, or `undefined` for anything else. */
function flattenPath(d: string): Point[][] | undefined {
	const tokens = d.match(/[A-Za-z]|-?\d*\.?\d+(?:e[-+]?\d+)?/giu);
	if (!tokens) {
		return undefined;
	}
	const subpaths: Point[][] = [];
	let current: Point[] = [];
	let pos: Point = [0, 0];
	let start: Point = [0, 0];
	let i = 0;
	let command = '';
	const num = (): number => Number.parseFloat(tokens[i++] ?? 'NaN');
	while (i < tokens.length) {
		const token = tokens[i] as string;
		if (/^[A-Za-z]$/u.test(token)) {
			command = token;
			i++;
			if (command === 'Z' || command === 'z') {
				if (current.length > 1) {
					subpaths.push(current);
				}
				current = [];
				pos = start;
				continue;
			}
		} else if (command === 'M') {
			command = 'L';
		}
		switch (command) {
			case 'M':
				if (current.length > 1) {
					subpaths.push(current);
				}
				pos = [num(), num()];
				start = pos;
				current = [pos];
				break;
			case 'L':
				pos = [num(), num()];
				current.push(pos);
				break;
			case 'H':
				pos = [num(), pos[1]];
				current.push(pos);
				break;
			case 'V':
				pos = [pos[0], num()];
				current.push(pos);
				break;
			case 'C': {
				const p1: Point = [num(), num()];
				const p2: Point = [num(), num()];
				const p3: Point = [num(), num()];
				for (let step = 1; step <= CURVE_STEPS; step++) {
					const t = step / CURVE_STEPS;
					const u = 1 - t;
					current.push([
						u * u * u * pos[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0],
						u * u * u * pos[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1],
					]);
				}
				pos = p3;
				break;
			}
			case 'Q': {
				const p1: Point = [num(), num()];
				const p2: Point = [num(), num()];
				for (let step = 1; step <= CURVE_STEPS; step++) {
					const t = step / CURVE_STEPS;
					const u = 1 - t;
					current.push([
						u * u * pos[0] + 2 * u * t * p1[0] + t * t * p2[0],
						u * u * pos[1] + 2 * u * t * p1[1] + t * t * p2[1],
					]);
				}
				pos = p2;
				break;
			}
			case 'A': {
				const rx = num();
				const ry = num();
				const rotation = num();
				const largeArc = num() !== 0;
				const sweep = num() !== 0;
				const to: Point = [num(), num()];
				current.push(...flattenArc(pos, rx, ry, rotation, largeArc, sweep, to));
				pos = to;
				break;
			}
			default:
				return undefined;
		}
	}
	if (current.length > 1) {
		subpaths.push(current);
	}
	return subpaths.every((sub) => sub.every(([x, y]) => Number.isFinite(x) && Number.isFinite(y)))
		? subpaths
		: undefined;
}

const fmt = (n: number): string => String(Math.round(n * 1000) / 1000);

/**
 * The clip a shape needs so a centred outline of `strokeWidth` shows in full:
 * the shape's own geometry plus a band of that width centred on its edge.
 *
 * @param clipPath    The resolved `path('...')` or `polygon(...)` clip.
 * @param strokeWidth Outline width in element pixels.
 * @param width       Element width in pixels.
 * @param height      Element height in pixels.
 * @returns A `path('...')` value, or the input unchanged when it cannot be widened.
 */
export function outsetClipPathForStroke(
	clipPath: string,
	strokeWidth: number,
	width: number,
	height: number,
): string {
	if (!(strokeWidth > 0)) {
		return clipPath;
	}
	const d = outlinePathData(clipPath, width, height);
	const subpaths = d ? flattenPath(d) : undefined;
	if (!d || !subpaths || subpaths.length === 0) {
		return clipPath;
	}
	const half = strokeWidth / 2;
	// Every quad must wind like the first subpath so the non-zero rule unions
	// them with it; a hole winds the other way and stays a hole.
	const signedArea = (pts: Point[]): number => {
		let area = 0;
		for (let k = 0; k < pts.length; k++) {
			const a = pts[k] as Point;
			const b = pts[(k + 1) % pts.length] as Point;
			area += a[0] * b[1] - b[0] * a[1];
		}
		return area;
	};
	const reference = Math.sign(signedArea(subpaths[0] as Point[])) || 1;
	const quads: string[] = [];
	for (const sub of subpaths) {
		const closed = [...sub, sub[0] as Point];
		for (let k = 0; k < closed.length - 1; k++) {
			const a = closed[k] as Point;
			const b = closed[k + 1] as Point;
			const len = Math.hypot(b[0] - a[0], b[1] - a[1]);
			if (len < 1e-6) {
				continue;
			}
			const nx = (-(b[1] - a[1]) / len) * half;
			const ny = ((b[0] - a[0]) / len) * half;
			let quad: Point[] = [
				[a[0] + nx, a[1] + ny],
				[b[0] + nx, b[1] + ny],
				[b[0] - nx, b[1] - ny],
				[a[0] - nx, a[1] - ny],
			];
			if (Math.sign(signedArea(quad)) !== reference) {
				quad = quad.reverse();
			}
			quads.push(`M ${quad.map(([x, y]) => `${fmt(x)} ${fmt(y)}`).join(' L ')} Z`);
		}
	}
	return `path('${d} ${quads.join(' ')}')`;
}
