import type { Point2 } from './svg-path-types';

/** Sample a cubic Bezier at `segments` steps (excluding the start point). */
export function sampleCubic(
	p0: Point2,
	p1: Point2,
	p2: Point2,
	p3: Point2,
	segments: number,
	out: Point2[],
): void {
	for (let i = 1; i <= segments; i++) {
		const t = i / segments;
		const mt = 1 - t;
		const x =
			mt * mt * mt * p0.x + 3 * mt * mt * t * p1.x + 3 * mt * t * t * p2.x + t * t * t * p3.x;
		const y =
			mt * mt * mt * p0.y + 3 * mt * mt * t * p1.y + 3 * mt * t * t * p2.y + t * t * t * p3.y;
		out.push({ x, y });
	}
}

/** Sample a quadratic Bezier at `segments` steps (excluding the start point). */
export function sampleQuadratic(
	p0: Point2,
	p1: Point2,
	p2: Point2,
	segments: number,
	out: Point2[],
): void {
	for (let i = 1; i <= segments; i++) {
		const t = i / segments;
		const mt = 1 - t;
		const x = mt * mt * p0.x + 2 * mt * t * p1.x + t * t * p2.x;
		const y = mt * mt * p0.y + 2 * mt * t * p1.y + t * t * p2.y;
		out.push({ x, y });
	}
}

/** An SVG elliptical arc in centre parameterization (SVG 1.1 appendix F.6.5). */
export interface SvgArcCenter {
	cx: number;
	cy: number;
	/** Radii, scaled up when too small to reach the end point (F.6.6). */
	rx: number;
	ry: number;
	/** The x-axis rotation in radians, with its cosine and sine. */
	phi: number;
	cosPhi: number;
	sinPhi: number;
	/** Start angle and signed sweep, in radians. */
	theta1: number;
	deltaTheta: number;
}

/**
 * The centre parameterization of an SVG arc from `start` to `end`, or
 * `undefined` when the arc is not drawn as one: a zero radius (a straight
 * line, as SVG draws it) or coincident end points (nothing).
 */
export function svgArcCenter(
	start: Point2,
	rxIn: number,
	ryIn: number,
	xAxisRotationDeg: number,
	largeArcFlag: boolean,
	sweepFlag: boolean,
	end: Point2,
): SvgArcCenter | undefined {
	if (rxIn === 0 || ryIn === 0 || (start.x === end.x && start.y === end.y)) {
		return undefined;
	}
	const phi = (xAxisRotationDeg * Math.PI) / 180;
	const cosPhi = Math.cos(phi);
	const sinPhi = Math.sin(phi);
	let rx = Math.abs(rxIn);
	let ry = Math.abs(ryIn);

	const dx2 = (start.x - end.x) / 2;
	const dy2 = (start.y - end.y) / 2;
	const x1p = cosPhi * dx2 + sinPhi * dy2;
	const y1p = -sinPhi * dx2 + cosPhi * dy2;

	const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
	if (lambda > 1) {
		const scale = Math.sqrt(lambda);
		rx *= scale;
		ry *= scale;
	}

	const sign = largeArcFlag !== sweepFlag ? 1 : -1;
	const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
	const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
	const coef = sign * Math.sqrt(Math.max(0, num / den));
	const cxp = (coef * (rx * y1p)) / ry;
	const cyp = (coef * -(ry * x1p)) / rx;

	const cx = cosPhi * cxp - sinPhi * cyp + (start.x + end.x) / 2;
	const cy = sinPhi * cxp + cosPhi * cyp + (start.y + end.y) / 2;

	const angle = (ux: number, uy: number, vx: number, vy: number): number => {
		const sign2 = ux * vy - uy * vx < 0 ? -1 : 1;
		const dot = Math.max(
			-1,
			Math.min(1, (ux * vx + uy * vy) / (Math.hypot(ux, uy) * Math.hypot(vx, vy))),
		);
		return sign2 * Math.acos(dot);
	};

	const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
	let deltaTheta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
	if (!sweepFlag && deltaTheta > 0) {
		deltaTheta -= 2 * Math.PI;
	} else if (sweepFlag && deltaTheta < 0) {
		deltaTheta += 2 * Math.PI;
	}
	return { cx, cy, rx, ry, phi, cosPhi, sinPhi, theta1, deltaTheta };
}

/** The point at angle `theta` of an arc's ellipse. */
export function svgArcPoint(arc: SvgArcCenter, theta: number): Point2 {
	const x = arc.rx * Math.cos(theta);
	const y = arc.ry * Math.sin(theta);
	return {
		x: arc.cx + x * arc.cosPhi - y * arc.sinPhi,
		y: arc.cy + x * arc.sinPhi + y * arc.cosPhi,
	};
}

/**
 * Sample an SVG elliptical arc (endpoint parameterization, SVG spec F.6) into
 * points, appended to `out`. Excludes the start point.
 */
export function sampleArc(
	start: Point2,
	rxIn: number,
	ryIn: number,
	xAxisRotationDeg: number,
	largeArcFlag: boolean,
	sweepFlag: boolean,
	end: Point2,
	segments: number,
	out: Point2[],
): void {
	const arc = svgArcCenter(start, rxIn, ryIn, xAxisRotationDeg, largeArcFlag, sweepFlag, end);
	if (!arc) {
		out.push(end);
		return;
	}
	const steps = Math.max(2, Math.round(segments * (Math.abs(arc.deltaTheta) / (Math.PI / 2))));
	for (let i = 1; i <= steps; i++) {
		out.push(svgArcPoint(arc, arc.theta1 + (arc.deltaTheta * i) / steps));
	}
}
