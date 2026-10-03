import type { Point2 } from './svg-path-types.js';

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
	if (rxIn === 0 || ryIn === 0 || (start.x === end.x && start.y === end.y)) {
		out.push(end);
		return;
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

	const steps = Math.max(2, Math.round(segments * (Math.abs(deltaTheta) / (Math.PI / 2))));
	for (let i = 1; i <= steps; i++) {
		const theta = theta1 + (deltaTheta * i) / steps;
		const x = cx + rx * Math.cos(theta) * cosPhi - ry * Math.sin(theta) * sinPhi;
		const y = cy + rx * Math.cos(theta) * sinPhi + ry * Math.sin(theta) * cosPhi;
		out.push({ x, y });
	}
}
