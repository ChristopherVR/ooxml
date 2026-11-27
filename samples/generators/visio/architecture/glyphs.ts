/**
 * The pictograms of the architecture masters, drawn here from lines, boxes and arcs in fractions
 * of a square: nothing is taken from another stencil.
 */
import { type Path, ellipse, polygon, polyline, rectangle } from '../lib/geometry';

const dot = (x: number, y: number, r = 0.045): Path => ellipse(x, y, r, r);
const stroke = (...points: [number, number][]): Path => polyline(points);

/** A mast with a louvred sensor housing and two anemometer cups. */
export const STATION: Path[] = [
	rectangle(0.3, 0.04, 0.7, 0.56),
	stroke([0.3, 0.2], [0.7, 0.2]),
	stroke([0.3, 0.38], [0.7, 0.38]),
	stroke([0.5, 0.56], [0.5, 0.84]),
	stroke([0.22, 0.84], [0.78, 0.84]),
	dot(0.14, 0.84, 0.09),
	dot(0.86, 0.84, 0.09),
];

/** A flat box with two antennas and status lights. */
export const HUB: Path[] = [
	rectangle(0.06, 0.1, 0.94, 0.5),
	stroke([0.24, 0.5], [0.14, 0.94]),
	stroke([0.76, 0.5], [0.86, 0.94]),
	dot(0.26, 0.3),
	dot(0.42, 0.3),
	dot(0.58, 0.3),
];

/** An octagon with traffic passing both ways. */
export const GATEWAY: Path[] = [
	polygon([
		[0.3, 0.02],
		[0.7, 0.02],
		[0.98, 0.3],
		[0.98, 0.7],
		[0.7, 0.98],
		[0.3, 0.98],
		[0.02, 0.7],
		[0.02, 0.3],
	]),
	stroke([0.24, 0.63], [0.76, 0.63]),
	stroke([0.63, 0.76], [0.76, 0.63], [0.63, 0.5]),
	stroke([0.76, 0.37], [0.24, 0.37]),
	stroke([0.37, 0.5], [0.24, 0.37], [0.37, 0.24]),
];

/** A pipe holding messages in order. */
export const QUEUE: Path[] = [
	rectangle(0.02, 0.24, 0.98, 0.76),
	stroke([0.26, 0.24], [0.26, 0.76]),
	stroke([0.5, 0.24], [0.5, 0.76]),
	stroke([0.74, 0.24], [0.74, 0.76]),
];

/** A hexagon with a smaller one inside: a deployable service. */
export const SERVICE: Path[] = [
	polygon([
		[0.26, 0.04],
		[0.74, 0.04],
		[0.98, 0.5],
		[0.74, 0.96],
		[0.26, 0.96],
		[0.02, 0.5],
	]),
	{
		...polygon([
			[0.39, 0.3],
			[0.61, 0.3],
			[0.72, 0.5],
			[0.61, 0.7],
			[0.39, 0.7],
			[0.28, 0.5],
		]),
		noFill: true,
	},
];

/** A cylinder with one band. */
export const DATABASE: Path[] = [
	{
		segments: [
			{ t: 'M', x: 0.12, y: 0.82 },
			{ t: 'L', x: 0.12, y: 0.18 },
			{ t: 'E', x: 0.88, y: 0.18, ax: 0.5, ay: 0.03, rx: 0.38, ry: 0.15 },
			{ t: 'L', x: 0.88, y: 0.82 },
			{ t: 'E', x: 0.12, y: 0.82, ax: 0.5, ay: 0.97, rx: 0.38, ry: 0.15 },
		],
	},
	{
		segments: [
			{ t: 'M', x: 0.12, y: 0.82 },
			{ t: 'E', x: 0.88, y: 0.82, ax: 0.5, ay: 0.67, rx: 0.38, ry: 0.15 },
		],
		noFill: true,
	},
	{
		segments: [
			{ t: 'M', x: 0.12, y: 0.52 },
			{ t: 'E', x: 0.88, y: 0.52, ax: 0.5, ay: 0.37, rx: 0.38, ry: 0.15 },
		],
		noFill: true,
	},
];

/** A storage box with a lid and a handle. */
export const ARCHIVE: Path[] = [
	rectangle(0.08, 0.06, 0.92, 0.66),
	rectangle(0.02, 0.66, 0.98, 0.92),
	stroke([0.38, 0.42], [0.62, 0.42]),
];

/** A browser window. */
export const WEB: Path[] = [
	rectangle(0.02, 0.1, 0.98, 0.9),
	stroke([0.02, 0.7], [0.98, 0.7]),
	dot(0.14, 0.8, 0.035),
	dot(0.26, 0.8, 0.035),
	dot(0.38, 0.8, 0.035),
];

/** A phone. */
export const MOBILE: Path[] = [
	rectangle(0.26, 0.02, 0.74, 0.98),
	stroke([0.26, 0.82], [0.74, 0.82]),
	stroke([0.26, 0.2], [0.74, 0.2]),
	dot(0.5, 0.11, 0.035),
];

/** A building: someone else's system. */
export const PARTNER: Path[] = [
	polygon([
		[0.1, 0.02],
		[0.9, 0.02],
		[0.9, 0.66],
		[0.5, 0.98],
		[0.1, 0.66],
	]),
	stroke([0.42, 0.02], [0.42, 0.32], [0.58, 0.32], [0.58, 0.02]),
	stroke([0.26, 0.5], [0.74, 0.5]),
];

/** A rack unit with three bays. */
export const SERVER: Path[] = [
	rectangle(0.12, 0.02, 0.88, 0.98),
	stroke([0.12, 0.34], [0.88, 0.34]),
	stroke([0.12, 0.66], [0.88, 0.66]),
	dot(0.26, 0.18, 0.04),
	dot(0.26, 0.5, 0.04),
	dot(0.26, 0.82, 0.04),
	stroke([0.42, 0.18], [0.74, 0.18]),
	stroke([0.42, 0.5], [0.74, 0.5]),
	stroke([0.42, 0.82], [0.74, 0.82]),
];

/** A wall of bricks. */
export const FIREWALL: Path[] = [
	rectangle(0.02, 0.1, 0.98, 0.9),
	stroke([0.02, 0.37], [0.98, 0.37]),
	stroke([0.02, 0.63], [0.98, 0.63]),
	stroke([0.5, 0.63], [0.5, 0.9]),
	stroke([0.26, 0.37], [0.26, 0.63]),
	stroke([0.74, 0.37], [0.74, 0.63]),
	stroke([0.5, 0.1], [0.5, 0.37]),
];

/** A circle that splits one path into two. */
export const BALANCER: Path[] = [
	ellipse(0.5, 0.5, 0.48, 0.48),
	stroke([0.16, 0.5], [0.46, 0.5]),
	stroke([0.46, 0.5], [0.78, 0.72]),
	stroke([0.46, 0.5], [0.78, 0.28]),
	stroke([0.64, 0.74], [0.78, 0.72], [0.72, 0.58]),
	stroke([0.72, 0.42], [0.78, 0.28], [0.64, 0.26]),
];

/** A cloud: three arcs over a flat base. */
export const CLOUD: Path[] = [
	{
		segments: [
			{ t: 'M', x: 0.24, y: 0.16 },
			{ t: 'L', x: 0.78, y: 0.16 },
			{ t: 'E', x: 0.76, y: 0.58, ax: 0.98, ay: 0.36, rx: 0.2, ry: 0.2 },
			{ t: 'E', x: 0.32, y: 0.62, ax: 0.56, ay: 0.9, rx: 0.25, ry: 0.25 },
			{ t: 'E', x: 0.24, y: 0.16, ax: 0.02, ay: 0.4, rx: 0.22, ry: 0.22 },
		],
	},
];
