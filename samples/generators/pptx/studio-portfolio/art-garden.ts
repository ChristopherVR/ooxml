/**
 * Case 02, Larkspur Heights Community Garden: three wayfinding totems in the garden, and the
 * pictogram set the signs are built from.
 */
import { Canvas, hex, type Rgb } from '../../lib/art';
import { arrow, drop, ellipse, gradientRect, leaf, shadow, sun, textLines } from './shapes';

const WHITE = hex('#FBF8F1');
const POST = hex('#2F4A3A');
const LARKSPUR = hex('#6A5BA6');
const MARIGOLD = hex('#E3A92F');
const POND = hex('#4E93B4');
const BRICK = hex('#C5553A');
const MOSS = hex('#56634F');

/** A seedling: a stalk with two leaves. */
function seedling(canvas: Canvas, cx: number, cy: number, size: number, color: Rgb): void {
	canvas.line(cx, cy + size * 0.6, cx, cy - size * 0.1, size * 0.1, color);
	leaf(canvas, cx + size * 0.3, cy - size * 0.28, size * 0.62, -0.6, color);
	leaf(canvas, cx - size * 0.3, cy - size * 0.28, size * 0.62, Math.PI + 0.6, color);
}

/** A tree: a round crown on a trunk. */
function tree(canvas: Canvas, cx: number, cy: number, size: number, color: Rgb): void {
	canvas.circle(cx, cy - size * 0.2, size * 0.5, color);
	canvas.rect(cx - size * 0.07, cy + size * 0.2, size * 0.14, size * 0.5, color);
}

/** A bench: a seat, a back and two legs. */
function bench(canvas: Canvas, cx: number, cy: number, size: number, color: Rgb): void {
	canvas.rect(cx - size * 0.62, cy - size * 0.36, size * 1.24, size * 0.16, color, 1, 3);
	canvas.rect(cx - size * 0.7, cy + size * 0.02, size * 1.4, size * 0.16, color, 1, 3);
	for (const side of [-1, 1])
		canvas.rect(
			cx + side * size * 0.48 - size * 0.06,
			cy + size * 0.18,
			size * 0.12,
			size * 0.38,
			color,
		);
}

/** A gate: an arch over two posts. */
function gate(canvas: Canvas, cx: number, cy: number, size: number, color: Rgb, ground: Rgb): void {
	canvas.circle(cx, cy - size * 0.05, size * 0.6, color);
	canvas.circle(cx, cy - size * 0.05, size * 0.4, ground);
	canvas.rect(cx - size * 0.62, cy - size * 0.05, size * 1.24, size * 0.7, ground);
	for (const side of [-1, 1])
		canvas.rect(
			cx + side * size * 0.5 - size * 0.1,
			cy - size * 0.05,
			size * 0.2,
			size * 0.66,
			color,
		);
}

type Pictogram = (canvas: Canvas, cx: number, cy: number, size: number, ground: Rgb) => void;

const PICTOGRAMS: [Rgb, Pictogram][] = [
	[MOSS, (c, x, y, s) => leaf(c, x, y + s * 0.08, s * 1.15, -Math.PI / 3, WHITE)],
	[POND, (c, x, y, s) => drop(c, x, y, s * 0.42, WHITE)],
	[MARIGOLD, (c, x, y, s) => sun(c, x, y, s * 0.58, WHITE)],
	[BRICK, (c, x, y, s) => arrow(c, x, y, s * 1.2, 0, s * 0.15, WHITE)],
	[LARKSPUR, (c, x, y, s) => seedling(c, x, y, s, WHITE)],
	[POST, (c, x, y, s) => tree(c, x, y, s, WHITE)],
	[hex('#8C7A5B'), (c, x, y, s) => bench(c, x, y, s * 0.8, WHITE)],
	[hex('#3D6E63'), (c, x, y, s, ground) => gate(c, x, y, s * 0.82, WHITE, ground)],
];

/** One totem: a post, a coloured head with a pictogram, a direction arrow and text bars. */
function totem(canvas: Canvas, x: number, top: number, width: number, head: number): void {
	const base = 520;
	const [color, pictogram] = PICTOGRAMS[head]!;
	shadow(canvas, x + width / 2 + 26, base + 4, width * 0.9, 13, 0.34, [28, 44, 34]);
	canvas.rect(x, top, width, base - top, POST, 1, 6);
	canvas.rect(x, top, width, width, color, 1, 6);
	canvas.rect(x, top + width - 8, width, 8, color);
	pictogram(canvas, x + width / 2, top + width / 2, width * 0.42, color);
	arrow(
		canvas,
		x + width * 0.3,
		top + width + 44,
		width * 0.34,
		head === 1 ? Math.PI : 0,
		7,
		WHITE,
	);
	textLines(
		canvas,
		x + width * 0.16,
		top + width + 78,
		[width * 0.62, width * 0.46],
		6,
		17,
		WHITE,
		0.9,
	);
	textLines(
		canvas,
		x + width * 0.16,
		top + width + 132,
		[width * 0.5, width * 0.58, width * 0.36],
		4,
		13,
		WHITE,
		0.55,
	);
}

export function gardenSigns(): Uint8Array {
	const canvas = new Canvas(960, 600, hex('#DCE5D2'));
	gradientRect(canvas, 0, 0, 960, 430, hex('#E6ECDD'), hex('#CBD8C0'));
	// Hedges in two tones behind the signs, then the path the signs stand on.
	for (const [cx, radius, tone] of [
		[90, 120, '#A9BA9B'],
		[300, 150, '#9DB08F'],
		[560, 130, '#A9BA9B'],
		[800, 160, '#9DB08F'],
		[960, 110, '#A9BA9B'],
	] as const)
		ellipse(canvas, cx, 440, radius * 1.5, radius, hex(tone));
	canvas.rect(0, 430, 960, 170, hex('#8FA083'));
	canvas.polygon(
		[
			[0, 560],
			[960, 470],
			[960, 600],
			[0, 600],
		],
		hex('#D9CDB6'),
	);
	totem(canvas, 150, 128, 132, 0);
	totem(canvas, 414, 68, 150, 4);
	totem(canvas, 706, 168, 118, 1);
	return canvas.toPng();
}

export function gardenPictograms(): Uint8Array {
	const ground = hex('#EFE9DD');
	const canvas = new Canvas(640, 400, ground);
	const size = 124;
	const gap = 24;
	const left = (640 - 4 * size - 3 * gap) / 2;
	const top = (400 - 2 * size - gap) / 2;
	PICTOGRAMS.forEach(([color, pictogram], index) => {
		const x = left + (index % 4) * (size + gap);
		const y = top + Math.floor(index / 4) * (size + gap);
		canvas.rect(x, y, size, size, color, 1, 14);
		pictogram(canvas, x + size / 2, y + size / 2, size * 0.42, color);
	});
	return canvas.toPng();
}
