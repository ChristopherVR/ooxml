/**
 * The keynote's backdrops, drawn pixel by pixel: the night gradient behind every slide, the dusk
 * sky of the cover and the closing slide, and the pressure map of the storm slide.
 */
import { Canvas, hex, mix, valueNoise, type Rgb } from '../../lib/art';
import { seeded } from '../../lib/random';
import { C } from './style';

/** An 8 x 8 Bayer matrix: ordered dither hides the banding of dark gradients and compresses well. */
const BAYER = [
	0, 32, 8, 40, 2, 34, 10, 42, 48, 16, 56, 24, 50, 18, 58, 26, 12, 44, 4, 36, 14, 46, 6, 38, 60, 28,
	52, 20, 62, 30, 54, 22, 3, 35, 11, 43, 1, 33, 9, 41, 51, 19, 59, 27, 49, 17, 57, 25, 15, 47, 7,
	39, 13, 45, 5, 37, 63, 31, 55, 23, 61, 29, 53, 21,
];

const clamp = (value: number) => Math.min(255, Math.max(0, value));
const smooth = (t: number) => {
	const k = Math.min(1, Math.max(0, t));
	return k * k * (3 - 2 * k);
};

/** Fills the canvas from `shade`, dithering the float result down to 8 bits. */
export function ditherFill(canvas: Canvas, shade: (u: number, v: number) => Rgb): Canvas {
	return canvas.fill((u, v) => {
		const x = Math.round(u * (canvas.width - 1));
		const y = Math.round(v * (canvas.height - 1));
		const t = (BAYER[(y & 7) * 8 + (x & 7)]! + 0.5) / 64 - 0.5;
		const color = shade(u, v);
		return [clamp(color[0] + t), clamp(color[1] + t), clamp(color[2] + t)];
	});
}

/** A colour ramp through evenly spaced stops. */
function ramp(stops: readonly Rgb[], t: number): Rgb {
	const k = Math.min(1, Math.max(0, t)) * (stops.length - 1);
	const index = Math.min(stops.length - 2, Math.floor(k));
	return mix(stops[index]!, stops[index + 1]!, k - index);
}

/** The backdrop of the slide master: deep navy falling to near black, lit from the upper left. */
export function backdrop(): Uint8Array {
	const top = hex('#111B30');
	const bottom = hex(C.night);
	const light = hex('#1C2B4D');
	return ditherFill(new Canvas(640, 360), (u, v) => {
		const base = mix(top, bottom, smooth(v * 1.1));
		const glow = Math.max(0, 1 - Math.hypot((u - 0.2) * 1.3, v - 0.05) / 0.75);
		const vignette = 1 - 0.35 * smooth(Math.hypot(u - 0.5, v - 0.5) * 1.3 - 0.35);
		const lit = mix(base, light, glow * glow * 0.55);
		return [lit[0] * vignette, lit[1] * vignette, lit[2] * vignette];
	}).toPng();
}

/** A ridgeline as a height (0 to 1 down the canvas) for each x. */
function ridge(seed: number, base: number, relief: number, scale: number) {
	const coarse = valueNoise(seed);
	const fine = valueNoise(seed + 17);
	return (u: number) =>
		base - relief * (coarse(u * scale, 0.5) - 0.5) * 2 - relief * 0.25 * fine(u * scale * 4, 0.5);
}

/**
 * Dusk over layered ridgelines, the last light low on the right, and on the nearest ridge a
 * roof with an Aerie mast on it. The left half stays dark for the headline.
 */
export function duskSky(): Uint8Array {
	const width = 960;
	const height = 540;
	const sky = [hex('#050814'), hex('#0B1030'), hex('#1F1B4F'), hex('#4A2E6E'), hex('#8A4A6A')];
	const ember = hex('#FFB36B');
	const sun = { u: 0.74, v: 0.66 };
	const canvas = ditherFill(new Canvas(width, height), (u, v) => {
		const base = ramp(sky, v / 0.72);
		const reach = Math.hypot((u - sun.u) * 1.15, (v - sun.v) * 2.4);
		const halo = Math.exp(-reach * reach * 5.5);
		const band = Math.exp(-(((v - sun.v) / 0.07) ** 2)) * Math.exp(-(((u - sun.u) / 0.42) ** 2));
		return mix(base, ember, Math.min(1, halo * 0.85 + band * 0.45));
	});
	const random = seeded(4101);
	for (let star = 0; star < 150; star++) {
		const x = random.range(0, width);
		const y = random.range(0, height * 0.5);
		const fade = 1 - y / (height * 0.5);
		canvas.circle(x, y, random.range(0.5, 1.2), hex('#DCE6FF'), random.range(0.15, 0.7) * fade);
	}
	// Four ridges, hazier and lighter the farther they are.
	const layers = [
		{ line: ridge(11, 0.66, 0.05, 2.2), color: hex('#3A2C5E'), haze: 0.45 },
		{ line: ridge(23, 0.72, 0.06, 3.1), color: hex('#231F47'), haze: 0.3 },
		{ line: ridge(37, 0.8, 0.07, 2.6), color: hex('#121530'), haze: 0.14 },
		{ line: ridge(53, 0.9, 0.06, 3.4), color: hex('#070A16'), haze: 0 },
	];
	for (const layer of layers) {
		canvas.paint(
			{ x0: 0, y0: 0, x1: width, y1: height },
			(x, y) => {
				const depth = (y / height - layer.line(x / width)) * 6;
				const warm = Math.exp(-(((x / width - sun.u) / 0.3) ** 2)) * layer.haze;
				return mix(mix(layer.color, ember, warm), hex(C.night), smooth(depth) * 0.6);
			},
			(x, y) => Math.min(1, Math.max(0, y - layer.line(x / width) * height + 0.5)),
		);
	}
	// The house on the nearest ridge: a roof, a chimney line, the mast and its lit ring.
	const houseX = 0.6 * width;
	const ground = layers[3]!.line(0.6) * height + 4;
	const dark = hex('#05070F');
	canvas.polygon(
		[
			[houseX - 46, ground],
			[houseX - 46, ground - 26],
			[houseX, ground - 54],
			[houseX + 46, ground - 26],
			[houseX + 46, ground],
		],
		dark,
	);
	canvas.rect(houseX + 22, ground - 96, 2.4, 60, dark);
	canvas.rect(houseX + 17, ground - 108, 12, 15, dark, 1, 5);
	canvas.circle(houseX + 23, ground - 100, 9, hex(C.amber), 0.22);
	canvas.rect(houseX + 17, ground - 102, 12, 2.4, hex(C.amber), 1, 1);
	canvas.rect(houseX - 24, ground - 20, 9, 9, hex('#FFCF8A'), 0.85);
	return canvas.toPng();
}

/**
 * The pressure map: land and sea, isobars closing around a low to the west, and the rain band
 * that the storm slide sets in motion. Positions are fractions of the picture.
 */
export function pressureMap(): Uint8Array {
	const width = 900;
	const height = 450;
	const terrain = valueNoise(907);
	const detail = valueNoise(911);
	const sea = hex('#0A1424');
	const land = hex('#131E33');
	const landAt = (u: number, v: number) =>
		terrain(u * 3.2, v * 2.4) * 0.7 + detail(u * 9, v * 8) * 0.3 + (u - 0.45) * 0.55;
	const pressure = (u: number, v: number) =>
		Math.hypot((u + 0.08) * 1.25, (v - 0.56) * 1.9) + (terrain(u * 1.6 + 5, v * 1.6) - 0.5) * 0.18;
	const canvas = ditherFill(new Canvas(width, height), (u, v) => {
		const shore = smooth((landAt(u, v) - 0.5) * 40);
		const relief = (detail(u * 14, v * 12) - 0.5) * 6 * shore;
		const base = mix(sea, land, shore);
		return [base[0] + relief, base[1] + relief, base[2] + relief];
	});
	const all = { x0: 0, y0: 0, x1: width, y1: height };
	// The coastline, then a faint graticule.
	canvas.paint(all, hex('#3A4B72'), (x, y) => {
		const edge = Math.abs(landAt(x / width, y / height) - 0.5) * 240;
		return Math.max(0, 1 - edge) * 0.7;
	});
	for (let x = 60; x < width; x += 120) canvas.rect(x, 0, 1, height, hex('#5B6C94'), 0.12);
	for (let y = 50; y < height; y += 100) canvas.rect(0, y, width, 1, hex('#5B6C94'), 0.12);
	// Isobars: contour lines of the pressure field, every tenth of its range.
	canvas.paint(all, hex('#8D9AB3'), (x, y) => {
		const value = pressure(x / width, y / height) * 9;
		const gradient =
			Math.hypot(
				pressure((x + 1) / width, y / height) - pressure((x - 1) / width, y / height),
				pressure(x / width, (y + 1) / height) - pressure(x / width, (y - 1) / height),
			) * 4.5;
		const distance = Math.abs(value - Math.round(value)) / Math.max(gradient, 1e-6);
		return Math.max(0, 1 - distance / 0.9) * 0.42;
	});
	return canvas.toPng();
}
