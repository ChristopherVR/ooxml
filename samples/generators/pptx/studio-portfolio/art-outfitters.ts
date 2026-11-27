/**
 * Case 03, Kestrelline Outfitters: the field guide's cover image (layered ridgelines at dusk,
 * with a faint contour overlay) and a spread from the book.
 */
import { Canvas, hex, mix, valueNoise, type Rgb } from '../../lib/art';
import { curve, shadow, textLines } from './shapes';

const NIGHT = hex('#161A26');
const DUSK = hex('#2B3A67');
const GLOW = hex('#E9A66B');
const CREAM = hex('#F4EEE2');
const VERMILION = hex('#E2452B');

/** The height of ridge `layer` at x: two octaves of smooth noise. */
function ridge(noise: (x: number, y: number) => number, layer: number, x: number): number {
	return noise(x / 190 + layer * 7.3, layer) * 0.72 + noise(x / 60 + layer * 3.1, layer + 9) * 0.28;
}

export function outfittersCover(): Uint8Array {
	const width = 960;
	const height = 600;
	const noise = valueNoise(41);
	const canvas = new Canvas(width, height, NIGHT);
	// Sky: night blue falling to a warm band at the horizon, banded so it compresses well.
	canvas.fill((_, v) => {
		const t = Math.round(Math.min(1, v / 0.62) * 22) / 22;
		return t < 0.7 ? mix(NIGHT, DUSK, t / 0.7) : mix(DUSK, GLOW, (t - 0.7) / 0.3);
	});
	// Contour rings, the map behind the mountains.
	for (let i = 1; i <= 7; i++) canvas.ring(190, 150, i * 34, 1.4, CREAM, 0.13);
	for (let i = 1; i <= 4; i++) canvas.ring(840, 96, i * 30, 1.4, CREAM, 0.1);
	// Six ridges from far and pale to near and dark.
	const layers = 6;
	for (let layer = 0; layer < layers; layer++) {
		const depth = layer / (layers - 1);
		const base = 286 + layer * 48;
		const swing = 70 + layer * 14;
		const color: Rgb = mix(hex('#7C86A8'), hex('#0E111B'), depth ** 0.8);
		canvas.paint({ x0: 0, y0: base - swing, x1: width, y1: height }, color, (x, y) =>
			Math.min(1, Math.max(0, y - (base - ridge(noise, layer, x) * swing) + 0.5)),
		);
	}
	// A dotted route climbing the nearest ridge.
	const near = layers - 1;
	for (let x = 452; x < 868; x += 16) {
		const y = 286 + near * 48 - ridge(noise, near, x) * (70 + near * 14) + 30;
		canvas.circle(x, y, 2.6, VERMILION);
	}
	return canvas.toPng();
}

export function outfittersSpread(): Uint8Array {
	const canvas = new Canvas(640, 400, hex('#C9CCD6'));
	canvas.rect(0, 300, 640, 100, hex('#BABECA'));
	shadow(canvas, 320, 338, 270, 20, 0.36, [20, 24, 40]);
	// Two pages and the gutter between them.
	canvas.rect(58, 62, 262, 270, CREAM, 1, 3);
	canvas.rect(320, 62, 262, 270, CREAM, 1, 3);
	for (let i = 0; i < 14; i++) canvas.rect(320 - i, 62, 1, 270, [60, 54, 46], 0.11 * (1 - i / 14));
	for (let i = 0; i < 14; i++) canvas.rect(320 + i, 62, 1, 270, [60, 54, 46], 0.08 * (1 - i / 14));
	// Left page: a contour map with the route in vermilion.
	canvas.rect(78, 82, 222, 166, hex('#DDE2D2'));
	const noise = valueNoise(7);
	for (let level = 1; level <= 6; level++)
		canvas.paint({ x0: 78, y0: 82, x1: 300, y1: 248 }, DUSK, (x, y) => {
			if (x < 78 || x > 300 || y < 82 || y > 248) return 0;
			const h = noise(x / 56, y / 56) * 0.7 + noise(x / 23, y / 23) * 0.3;
			return Math.abs(h - level / 7) < 0.006 ? 0.55 : 0;
		});
	curve(
		canvas,
		[
			[92, 232],
			[128, 206],
			[150, 176],
			[196, 160],
			[226, 126],
			[282, 100],
		],
		3,
		VERMILION,
	);
	canvas.circle(282, 100, 6, VERMILION);
	textLines(canvas, 78, 266, [150, 196, 120], 4, 12, hex('#5B564E'), 0.7);
	// Right page: a heading, body text and the route's elevation profile.
	canvas.rect(340, 84, 132, 12, NIGHT, 1, 2);
	canvas.rect(340, 106, 54, 5, VERMILION);
	textLines(canvas, 340, 130, [220, 206, 222, 188, 214, 120], 4, 12, hex('#5B564E'), 0.7);
	const profile: [number, number][] = [[340, 312]];
	for (let i = 0; i <= 22; i++) profile.push([340 + i * 10, 300 - noise(i / 5, 3) * 70 - i * 1.2]);
	profile.push([560, 312]);
	canvas.polygon(profile, DUSK, 0.9);
	return canvas.toPng();
}
