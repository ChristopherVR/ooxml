/**
 * The two generated pictures of the deck: the network map of Port Averly (land, harbour, tram,
 * bus and ferry lines, stops) and the route-diagram panel on the cover. The map's geometry is
 * exported in map units so the slides can place labels and the moving tram exactly on it.
 */
import { Canvas, hex, mix, type Rgb } from '../../lib/art';
import { seeded } from '../../lib/random';
import { C } from './style';

type Point = readonly [number, number];

/** Map units: 1000 x 580. */
export const MAP_W = 1000;
export const MAP_H = 580;

/** T1 Harbour Line, Marrow Hill to the old terminus at Customs House. */
export const T1: Point[] = [
	[70, 440],
	[250, 404],
	[430, 344],
	[600, 330],
	[700, 296],
	[790, 246],
];
/** The 2025 extension along the north shore, with its four new stops. */
export const T1_EXTENSION: Point[] = [
	[790, 246],
	[836, 196],
	[878, 146],
	[914, 96],
	[946, 44],
];
/** T2 Northgate Line. */
export const T2: Point[] = [
	[110, 80],
	[290, 146],
	[460, 246],
	[600, 330],
	[688, 366],
];
const FERRY: Point[][] = [
	[
		[688, 372],
		[720, 440],
		[672, 486],
		[604, 492],
	],
	[
		[688, 372],
		[790, 380],
		[880, 330],
		[930, 292],
	],
];
const BUSES: Point[][] = [
	[
		[600, 330],
		[470, 420],
		[330, 470],
		[150, 530],
	],
	[
		[40, 250],
		[220, 270],
		[430, 344],
		[560, 440],
		[604, 492],
	],
	[
		[290, 146],
		[330, 300],
		[330, 470],
	],
	[
		[460, 246],
		[640, 190],
		[836, 196],
	],
	[
		[520, 40],
		[560, 180],
		[600, 330],
	],
];

function smooth(points: Point[], passes: number): Point[] {
	let out = points;
	for (let pass = 0; pass < passes; pass++) {
		const next: Point[] = [out[0]!];
		for (let i = 0; i < out.length - 1; i++) {
			const [ax, ay] = out[i]!;
			const [bx, by] = out[i + 1]!;
			next.push([ax * 0.75 + bx * 0.25, ay * 0.75 + by * 0.25]);
			next.push([ax * 0.25 + bx * 0.75, ay * 0.25 + by * 0.75]);
		}
		next.push(out[out.length - 1]!);
		out = next;
	}
	return out;
}

function polyline(canvas: Canvas, points: Point[], k: number, width: number, color: Rgb): void {
	for (let i = 0; i < points.length - 1; i++)
		canvas.line(
			points[i]![0] * k,
			points[i]![1] * k,
			points[i + 1]![0] * k,
			points[i + 1]![1] * k,
			width,
			color,
		);
}

function dashed(canvas: Canvas, points: Point[], k: number, width: number, color: Rgb): void {
	const fine = smooth(points, 3);
	let travelled = 0;
	for (let i = 0; i < fine.length - 1; i++) {
		const [ax, ay] = fine[i]!;
		const [bx, by] = fine[i + 1]!;
		const length = Math.hypot(bx - ax, by - ay);
		if (Math.floor(travelled / 14) % 2 === 0)
			canvas.line(ax * k, ay * k, bx * k, by * k, width, color);
		travelled += length;
	}
}

function stop(canvas: Canvas, point: Point, k: number, radius: number, color: Rgb): void {
	canvas.circle(point[0] * k, point[1] * k, radius * k, color);
	canvas.circle(point[0] * k, point[1] * k, radius * k * 0.55, hex(C.white));
}

/** The network map, `width` px wide. */
export function networkMap(width = 912): Uint8Array {
	const k = width / MAP_W;
	const land = hex('#F3EEE2');
	const canvas = new Canvas(width, Math.round(MAP_H * k), land);
	const random = seeded(77);
	// A faint, slightly skewed street grid.
	const street = hex('#E8E1D1');
	for (let i = -8; i < 30; i++) {
		canvas.line(i * 44 * k, 0, (i * 44 + 120) * k, MAP_H * k, 1, street);
		canvas.line(0, (i * 40 + 30) * k, MAP_W * k, (i * 40 - 60) * k, 1, street);
	}
	// Parks.
	const park = hex('#DDE6D3');
	for (const [x, y, r] of [
		[200, 210, 46],
		[520, 120, 38],
		[380, 520, 52],
		[730, 130, 30],
	] as const)
		for (let blob = 0; blob < 5; blob++)
			canvas.circle(
				(x + random.range(-r * 0.5, r * 0.5)) * k,
				(y + random.range(-r * 0.4, r * 0.4)) * k,
				r * k * random.range(0.5, 0.8),
				park,
			);
	// The harbour: shallow water at the shore, deeper further out.
	const coast = smooth(
		[
			[1000, 262],
			[930, 300],
			[850, 340],
			[760, 372],
			[700, 384],
			[650, 430],
			[610, 500],
			[560, 540],
			[530, 580],
		],
		3,
	);
	const shallow = hex('#CFE4E4');
	const deep = hex('#B4D5D7');
	const water = [...coast, [1000, 580] as Point];
	canvas.polygon(
		water.map(([x, y]) => [x * k, y * k] as const),
		shallow,
	);
	canvas.polygon(
		water.map(([x, y]) => [(x + 46) * k, (y + 52) * k] as const),
		mix(shallow, deep, 0.6),
	);
	canvas.polygon(
		water.map(([x, y]) => [(x + 110) * k, (y + 120) * k] as const),
		deep,
	);
	polyline(canvas, coast, k, 2, hex('#FFFFFF'));
	// Lines: buses thin, ferries dashed, trams bold.
	for (const bus of BUSES) polyline(canvas, smooth(bus, 2), k, 2.2, hex(C.tealSoft));
	for (const ferry of FERRY) dashed(canvas, ferry, k, 2.6, hex(C.ochre));
	polyline(canvas, T2, k, 6, hex(C.teal));
	polyline(canvas, T1, k, 6, hex(C.coral));
	polyline(canvas, T1_EXTENSION, k, 6, hex(C.coral));
	for (const point of T2) stop(canvas, point, k, 8, hex(C.teal));
	for (const point of T1) stop(canvas, point, k, 8, hex(C.coral));
	for (const point of T1_EXTENSION.slice(1)) {
		canvas.ring(point[0] * k, point[1] * k, 15 * k, 2, hex(C.coral), 0.55);
		stop(canvas, point, k, 9, hex(C.coral));
	}
	for (const pier of [FERRY[0]![3]!, FERRY[1]![3]!]) stop(canvas, pier, k, 7, hex(C.ochre));
	// The interchange where both tram lines meet.
	canvas.circle(600 * k, 330 * k, 12 * k, hex(C.ink));
	canvas.circle(600 * k, 330 * k, 7 * k, hex(C.white));
	return canvas.toPng();
}

/** The cover panel: a transit diagram in paper lines on harbour teal. */
export function coverPanel(): Uint8Array {
	const w = 500;
	const h = 720;
	const top = hex('#0B6A73');
	const bottom = hex('#0E7C86');
	const canvas = new Canvas(w, h).fill((_, v) => mix(top, bottom, v));
	const paper = hex(C.paper);
	const soft = mix(bottom, paper, 0.35);
	const route = (points: Point[], width: number, color: Rgb) => {
		polyline(canvas, points, 1, width, color);
		for (const point of points) {
			canvas.circle(point[0], point[1], width * 1.25, color);
			canvas.circle(point[0], point[1], width * 0.62, bottom);
		}
	};
	route(
		[
			[-40, 150],
			[110, 150],
			[250, 290],
			[250, 470],
			[390, 610],
			[540, 610],
		],
		9,
		soft,
	);
	route(
		[
			[70, -40],
			[70, 240],
			[250, 420],
			[430, 420],
			[540, 310],
		],
		9,
		paper,
	);
	route(
		[
			[-40, 560],
			[130, 560],
			[250, 440],
			[380, 310],
			[380, 120],
			[540, 120],
		],
		9,
		hex(C.coral),
	);
	// The interchange.
	canvas.rect(222, 396, 56, 92, paper, 1, 28);
	canvas.rect(232, 406, 36, 72, bottom, 1, 18);
	return canvas.toPng();
}

/** A point of the map in slide coordinates, for a map drawn at `box`. */
export function onSlide(
	point: Point,
	box: { x: number; y: number; width: number },
): { x: number; y: number } {
	const k = box.width / MAP_W;
	return { x: box.x + point[0] * k, y: box.y + point[1] * k };
}
