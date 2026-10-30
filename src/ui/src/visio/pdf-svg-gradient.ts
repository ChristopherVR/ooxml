/**
 * SVG `linearGradient` and `radialGradient` elements as PDF shadings. A shading is painted with
 * `sh` inside the filled path's clip, so it needs no pattern matrix: the matrices below are
 * applied with `cm` first (bounding-box units, then `gradientTransform`).
 */
import type { PdfShading } from '../export/pdf-vector';
import type { PathSegment } from './pdf-path';
import { number, parseColor, parseTransform } from './pdf-svg-style';

export interface GradientPaint {
	shading: PdfShading;
	/** Matrices to concatenate, in order, before painting the shading. */
	matrices: number[][];
	/** Mean stop opacity: a shading carries no per-stop alpha. */
	opacity: number;
	/** Whether the stop opacities differed (so the caller can say it averaged them). */
	variableOpacity: boolean;
	/** `repeat` and `reflect` spreads are painted as `pad`. */
	spreadApproximated: boolean;
}

export const isGradient = (element: Element | undefined): element is Element =>
	element?.localName === 'linearGradient' || element?.localName === 'radialGradient';

/** Bounds of a path's points (control points included). */
export function pathBounds(
	segments: readonly PathSegment[],
): [number, number, number, number] | undefined {
	let minX = Infinity,
		minY = Infinity,
		maxX = -Infinity,
		maxY = -Infinity;
	const take = (x: number, y: number) => {
		minX = Math.min(minX, x);
		minY = Math.min(minY, y);
		maxX = Math.max(maxX, x);
		maxY = Math.max(maxY, y);
	};
	for (const segment of segments) {
		if (segment.t === 'M' || segment.t === 'L') take(segment.x, segment.y);
		else if (segment.t === 'C') {
			take(segment.x1, segment.y1);
			take(segment.x2, segment.y2);
			take(segment.x, segment.y);
		}
	}
	return Number.isFinite(minX) ? [minX, minY, maxX - minX, maxY - minY] : undefined;
}

function length(value: string | null, fallback: number, bbox: boolean): number {
	if (value === null) return fallback;
	const parsed = Number.parseFloat(value);
	if (!Number.isFinite(parsed)) return fallback;
	return value.trim().endsWith('%') && bbox ? parsed / 100 : parsed;
}

/** The gradient as a shading for a path with `bounds`; `undefined` when it cannot be drawn. */
export function gradientPaint(
	gradient: Element,
	bounds: [number, number, number, number] | undefined,
): GradientPaint | undefined {
	const units = gradient.getAttribute('gradientUnits') !== 'userSpaceOnUse';
	const stops = Array.from(gradient.children)
		.filter((child) => child.localName === 'stop')
		.map((stop) => {
			const raw = stop.getAttribute('offset') ?? '0';
			const offset = raw.trim().endsWith('%') ? number(raw) / 100 : number(raw);
			const color = parseColor(stop.getAttribute('stop-color') ?? '#000000') ?? {
				r: 0,
				g: 0,
				b: 0,
				a: 0,
			};
			return {
				offset: Math.min(1, Math.max(0, offset)),
				r: color.r,
				g: color.g,
				b: color.b,
				a: color.a * number(stop.getAttribute('stop-opacity'), 1),
			};
		})
		.sort((a, b) => a.offset - b.offset);
	if (!stops.length) return undefined;
	if (stops.length === 1) stops.push({ ...stops[0]!, offset: 1 });
	const matrices: number[][] = [];
	if (units) {
		if (!bounds || !(bounds[2] > 0) || !(bounds[3] > 0)) return undefined;
		matrices.push([bounds[2], 0, 0, bounds[3], bounds[0], bounds[1]]);
	}
	matrices.push(...parseTransform(gradient.getAttribute('gradientTransform')));
	const opacities = stops.map((stop) => stop.a);
	const shading: PdfShading =
		gradient.localName === 'linearGradient'
			? {
					kind: 'axial',
					coords: [
						length(gradient.getAttribute('x1'), 0, units),
						length(gradient.getAttribute('y1'), 0, units),
						length(gradient.getAttribute('x2'), units ? 1 : 0, units),
						length(gradient.getAttribute('y2'), 0, units),
					],
					stops,
				}
			: (() => {
					const cx = length(gradient.getAttribute('cx'), units ? 0.5 : 0, units);
					const cy = length(gradient.getAttribute('cy'), units ? 0.5 : 0, units);
					const r = length(gradient.getAttribute('r'), units ? 0.5 : 0, units);
					return {
						kind: 'radial' as const,
						coords: [
							length(gradient.getAttribute('fx'), cx, units),
							length(gradient.getAttribute('fy'), cy, units),
							length(gradient.getAttribute('fr'), 0, units),
							cx,
							cy,
							r,
						],
						stops,
					};
				})();
	if (shading.kind === 'radial' && !(shading.coords[5]! > 0)) return undefined;
	const spread = gradient.getAttribute('spreadMethod');
	return {
		shading,
		matrices,
		opacity: opacities.reduce((sum, value) => sum + value, 0) / opacities.length,
		variableOpacity: opacities.some((value) => Math.abs(value - opacities[0]!) > 0.01),
		spreadApproximated: spread === 'repeat' || spread === 'reflect',
	};
}
