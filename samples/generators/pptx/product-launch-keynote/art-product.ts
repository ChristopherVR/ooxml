/**
 * The product pictures: the Aerie station as a cut-out render (colour plus an alpha plane, so it
 * sits on any slide and can Morph across them) and the screen of the Aerie app.
 */
import { encodePng } from '../../../../src/core/drawingml/raster-encode';
import { Canvas, hex, mix, type Rgb } from '../../lib/art';
import { seeded } from '../../lib/random';
import { ditherFill } from './art-scenes';
import { C } from './style';

type Shade = (x: number, y: number) => Rgb;
type Area = readonly [x: number, y: number, w: number, h: number];
const WHITE: Rgb = [255, 255, 255];
const unit = (value: number) => Math.min(1, Math.max(0, value));

/** A colour canvas and a coverage canvas that are always drawn together. */
class Sprite {
	private readonly color: Canvas;
	private readonly mask: Canvas;

	constructor(
		readonly width: number,
		readonly height: number,
		matte: Rgb,
	) {
		this.color = new Canvas(width, height, matte);
		this.mask = new Canvas(width, height, [0, 0, 0]);
	}

	/** Blends `color` wherever `coverage` is positive, in both planes, inside `box` (x, y, w, h). */
	paint(
		color: Rgb | Shade,
		coverage: (x: number, y: number) => number,
		[x, y, w, h]: Area = [0, 0, this.width, this.height],
	): this {
		const box = { x0: x, y0: y, x1: x + w, y1: y + h };
		this.color.paint(box, color, coverage);
		this.mask.paint(box, WHITE, coverage);
		return this;
	}

	line(x0: number, y0: number, x1: number, y1: number, stroke: number, color: Rgb): this {
		this.color.line(x0, y0, x1, y1, stroke, color);
		this.mask.line(x0, y0, x1, y1, stroke, WHITE);
		return this;
	}

	circle(cx: number, cy: number, radius: number, color: Rgb | Shade, alpha = 1): this {
		return this.paint(color, (x, y) => alpha * unit(radius - Math.hypot(x - cx, y - cy) + 0.5), [
			cx - radius - 1,
			cy - radius - 1,
			2 * radius + 2,
			2 * radius + 2,
		]);
	}

	toPng(): Uint8Array {
		const alpha = new Uint8Array(this.width * this.height);
		for (let index = 0; index < alpha.length; index++) alpha[index] = this.mask.pixels[index * 3]!;
		return encodePng(this.width, this.height, this.color.pixels, alpha);
	}
}

/** Coverage of a rounded rectangle (antialiased). */
function roundedBox(x: number, y: number, w: number, h: number, radius: number) {
	return (px: number, py: number) => {
		const dx = Math.max(x + radius - px, 0, px - (x + w - radius));
		const dy = Math.max(y + radius - py, 0, py - (y + h - radius));
		if (px < x || px > x + w || py < y || py > y + h) return 0;
		return dx === 0 && dy === 0 ? 1 : unit(radius - Math.hypot(dx, dy) + 0.5);
	};
}

export const DEVICE_SIZE = { width: 480, height: 660 } as const;

/**
 * The Aerie station: a louvred pod with a lit sensor ring, wind cups on top and a small solar
 * panel on the mast, lit from the upper left with a cool rim light on the right.
 */
export function deviceRender(): Uint8Array {
	const { width, height } = DEVICE_SIZE;
	const sprite = new Sprite(width, height, hex('#0A0F1C'));
	const cx = 240;
	const amber = hex(C.amber);
	const steelLight = hex('#E4EBF7');
	const steelDark = hex('#3A4763');
	// The bloom of the ring, behind everything.
	sprite.paint(amber, (x, y) => {
		const reach = Math.hypot((x - cx) / 215, (y - 276) / 120);
		return 0.42 * Math.exp(-reach * reach * 3.2);
	});
	// The mast, fading out towards the bottom of the picture.
	sprite.paint(
		(x) => mix(steelLight, steelDark, unit((x - (cx - 9)) / 18) ** 0.7),
		(x, y) => roundedBox(cx - 9, 370, 18, 330, 2)(x, y) * unit((height - y) / 150),
		[cx - 10, 368, 20, 292],
	);
	// The solar panel on its bracket.
	sprite.line(cx - 8, 470, cx - 62, 452, 5, hex('#55627F'));
	const panel = (x: number, y: number) => {
		const skew = (y - 418) * 0.32;
		return roundedBox(96 + skew, 418, 96, 62, 5)(x, y);
	};
	const panelArea: Area = [90, 414, 130, 70];
	sprite.paint(
		(x, y) => mix(hex('#22355F'), hex('#0E1830'), unit((x - 90 + (y - 418)) / 150)),
		panel,
		panelArea,
	);
	sprite.paint(
		hex('#6F8FD0'),
		(x, y) => (panel(x, y) > 0 && (y - 418) % 21 < 1.2 ? 0.5 : 0),
		panelArea,
	);
	// The wind cups: a post, three arms, three cups.
	sprite.line(cx, 122, cx, 62, 5, hex('#AEB9CE'));
	for (const [dx, dy, r] of [
		[-78, -4, 15],
		[72, -10, 13],
		[10, 8, 17],
	] as const) {
		sprite.line(cx, 62, cx + dx, 62 + dy, 3.5, hex('#8C99B3'));
		sprite.circle(cx + dx, 62 + dy, r, (x, y) =>
			mix(steelLight, steelDark, unit(0.5 + (x - (cx + dx)) / (2.2 * r) + (y - 62 - dy) / (3 * r))),
		);
	}
	sprite.circle(cx, 62, 7, hex('#D5DDEB'));
	// The pod: a capsule shaded across its width, darker towards the base.
	const pod = roundedBox(cx - 108, 118, 216, 276, 104);
	const podArea: Area = [cx - 112, 114, 224, 284];
	sprite.paint(
		(x, y) => {
			const across = unit((x - (cx - 108)) / 216);
			const curve = unit(0.08 + 0.92 * Math.abs(across - 0.3) ** 0.9 * 1.35);
			const down = unit((y - 118) / 276);
			const shaded = mix(steelLight, steelDark, unit(curve + down * 0.22));
			const gloss = Math.exp(-(((across - 0.26) / 0.07) ** 2)) * (1 - down) * 0.5;
			return mix(shaded, WHITE, gloss);
		},
		pod,
		podArea,
	);
	// The cool rim light along the right edge.
	sprite.paint(
		hex(C.sky),
		(x, y) => {
			const inside = pod(x, y);
			return inside * (1 - pod(x + 5, y - 1)) * 0.85;
		},
		podArea,
	);
	// The louvres of the lower half.
	for (const y0 of [312, 332, 352, 372])
		sprite.paint(
			hex('#141C2E'),
			(x, y) => pod(x, y) * roundedBox(cx - 120, y0, 240, 5, 2)(x, y) * 0.55,
			podArea,
		);
	// The sensor ring: a dark groove with the lit core inside it.
	sprite.paint(
		hex('#141C2E'),
		(x, y) => pod(x, y) * roundedBox(0, 262, width, 28, 0)(x, y) * 0.92,
		podArea,
	);
	sprite.paint(
		(x) => mix(hex('#FFF1D2'), amber, unit(Math.abs(x - (cx - 30)) / 120)),
		(x, y) => pod(x, y) * roundedBox(0, 271, width, 10, 0)(x, y),
		podArea,
	);
	// The ring lights the shell above and below it.
	sprite.paint(amber, (x, y) => pod(x, y) * 0.38 * Math.exp(-(((y - 276) / 30) ** 2)), podArea);
	sprite.circle(cx + 56, 222, 3.5, hex(C.sky));
	return sprite.toPng();
}

export const APP_SIZE = { width: 360, height: 740 } as const;

/** Where the slide lays its live text over the app screen, in picture pixels. */
export const APP_ZONES = {
	chart: { top: 276, bottom: 404 },
	rain: { top: 440, bottom: 500 },
	cards: [
		{ top: 528, bottom: 604 },
		{ top: 616, bottom: 692 },
	],
} as const;

/** The home screen of the Aerie app, without its text (the slide sets that in real type). */
export function appScreen(): Uint8Array {
	const { width, height } = APP_SIZE;
	const top = hex('#13203C');
	const bottom = hex('#0A101E');
	const canvas = ditherFill(new Canvas(width, height), (u, v) => {
		const glow = Math.exp(-(((u - 0.86) / 0.4) ** 2 + ((v - 0.1) / 0.16) ** 2));
		return mix(mix(top, bottom, v), hex('#6B4A2A'), glow * 0.45);
	});
	const sky = hex(C.sky);
	const amber = hex(C.amber);
	// Temperature over the next 48 hours: a filled curve with a "now" marker.
	const { top: chartTop, bottom: chartBottom } = APP_ZONES.chart;
	const curve = (x: number) => {
		const t = (x - 24) / (width - 48);
		return (
			chartTop + 46 - 30 * Math.sin(t * Math.PI * 4 - 0.6) * (1 - t * 0.35) + 16 * Math.sin(t * 9)
		);
	};
	canvas.paint({ x0: 24, y0: chartTop, x1: width - 24, y1: chartBottom }, sky, (x, y) =>
		y > curve(x) ? 0.3 * unit(1 - (y - curve(x)) / (chartBottom - curve(x))) : 0,
	);
	for (let x = 24; x < width - 24; x += 2) canvas.line(x, curve(x), x + 2, curve(x + 2), 2.6, sky);
	canvas.circle(96, curve(96), 9, amber, 0.3);
	canvas.circle(96, curve(96), 4.5, amber);
	canvas.rect(24, chartBottom, width - 48, 1, hex('#2A3654'));
	// Rain per hour: quiet, then the shower the app is warning about.
	const random = seeded(7702);
	const rainBottom = APP_ZONES.rain.bottom;
	for (let bar = 0; bar < 16; bar++) {
		const wet = bar >= 6 && bar <= 10;
		const size = wet ? [22, 44, 58, 40, 18][bar - 6]! : random.range(3, 8);
		canvas.rect(26 + bar * 19.6, rainBottom - size, 12, size, wet ? sky : hex('#3A4763'), 1, 3);
	}
	canvas.rect(24, rainBottom + 5, width - 48, 1, hex('#2A3654'));
	// Two cards with a status dot each, and the tab bar.
	APP_ZONES.cards.forEach((card, index) => {
		canvas.rect(20, card.top, width - 40, card.bottom - card.top, hex('#18243F'), 1, 14);
		canvas.circle(50, (card.top + card.bottom) / 2, 13, index === 0 ? amber : sky, 0.22);
		canvas.circle(50, (card.top + card.bottom) / 2, 5.5, index === 0 ? amber : sky);
	});
	canvas.rect(0, 706, width, 1, hex('#22304F'));
	for (let tab = 0; tab < 4; tab++)
		canvas.circle(64 + tab * 77.3, 723, 4, tab === 0 ? amber : hex('#4A5878'));
	return canvas.toPng();
}
