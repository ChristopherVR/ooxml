/** The newsletter's pictures, drawn pixel by pixel in the garden palette. */
import { Canvas, hex, mix, type Rgb } from '../../lib/art';
import { seeded } from '../../lib/random';
import { GARDEN } from './styles';

const tone = (name: keyof typeof GARDEN): Rgb => hex(GARDEN[name]);

/** A flower head: a ring of round petals about a centre. */
function flower(canvas: Canvas, x: number, y: number, size: number, petal: Rgb, heart: Rgb): void {
	const petals = 12;
	for (let index = 0; index < petals; index++) {
		const angle = (index / petals) * Math.PI * 2;
		canvas.circle(
			x + Math.cos(angle) * size * 0.72,
			y + Math.sin(angle) * size * 0.72,
			size * 0.3,
			petal,
		);
	}
	canvas.circle(x, y, size * 0.52, heart);
}

/** The masthead band: a forest ground with leaves and flower heads gathering to the right. */
export function masthead(): Uint8Array {
	const width = 1240;
	const height = 340;
	const canvas = new Canvas(width, height, tone('forest'));
	const random = seeded(4107);
	const greens: Rgb[] = [
		mix(tone('forest'), tone('leaf'), 0.35),
		mix(tone('forest'), tone('leaf'), 0.7),
		tone('leaf'),
		mix(tone('forest'), [0, 0, 0], 0.18),
	];
	// Leaves: pairs of discs along stems, denser towards the right so the title stays clear.
	for (let stem = 0; stem < 26; stem++) {
		const u = 0.5 + 0.5 * random.next() ** 0.75;
		const x = u * width;
		const top = height * random.range(0.05, 0.75);
		const lean = random.range(-40, 40);
		const green = random.pick(greens);
		canvas.line(x, height + 10, x + lean, top, 3, green);
		const leaves = random.int(3, 6);
		for (let leaf = 0; leaf < leaves; leaf++) {
			const k = (leaf + 1) / (leaves + 1);
			const lx = x + lean * (1 - k);
			const ly = top + (height - top) * k;
			const side = leaf % 2 ? 1 : -1;
			canvas.circle(lx + side * 13, ly - 4, random.range(9, 15), green);
		}
	}
	const heads: [number, number, number, Rgb, Rgb][] = [
		[0.9, 0.34, 64, tone('marigold'), tone('bark')],
		[0.74, 0.62, 40, tone('terracotta'), tone('marigold')],
		[0.6, 0.3, 30, tone('larkspur'), tone('cream')],
		[0.985, 0.8, 46, tone('cream'), tone('marigold')],
		[0.665, 0.87, 22, tone('marigold'), tone('terracotta')],
	];
	for (const [u, v, size, petal, heart] of heads)
		flower(canvas, u * width, v * height, size, petal, heart);
	return canvas.toPng();
}

/** The new raised beds seen from above: paths, twelve beds in rows and the hedge along the top. */
export function bedsPlan(): Uint8Array {
	const width = 760;
	const height = 520;
	const grass = hex('A9C98A');
	const canvas = new Canvas(width, height, grass);
	const random = seeded(88);
	// The hedge and the gravel paths.
	canvas.rect(0, 0, width, 54, tone('forest'));
	for (let x = 14; x < width; x += 30)
		canvas.circle(
			x + random.range(-4, 4),
			28 + random.range(-6, 6),
			19,
			mix(tone('forest'), tone('leaf'), 0.4),
		);
	const path = tone('cream');
	canvas.rect(0, 250, width, 34, path);
	canvas.rect(363, 54, 34, height - 54, path);
	// Beds: soil with rows of seedlings, three colours of crop.
	const crops: Rgb[] = [
		tone('leaf'),
		mix(tone('leaf'), tone('marigold'), 0.45),
		mix(tone('forest'), tone('leaf'), 0.5),
	];
	const soil = hex('7A5A3C');
	for (let row = 0; row < 2; row++)
		for (let column = 0; column < 6; column++) {
			const x = 22 + column * 113 + (column >= 3 ? 46 : 0);
			const y = 78 + row * 224;
			canvas.rect(x - 3, y - 3, 100, 160, tone('bark'), 1, 8);
			canvas.rect(x, y, 94, 154, soil, 1, 6);
			const crop = crops[(row * 2 + column) % crops.length]!;
			for (let line = 0; line < 3; line++)
				for (let plant = 0; plant < 7; plant++)
					canvas.circle(x + 18 + line * 29, y + 16 + plant * 20.5, random.range(5.5, 8), crop);
		}
	// The water butt and the tool shed by the gate.
	canvas.circle(690, 492, 17, tone('sky'));
	canvas.ring(690, 492, 17, 3, tone('ink'), 0.5);
	canvas.rect(28, 470, 86, 40, tone('terracotta'), 1, 4);
	return canvas.toPng();
}

/** A sunflower on white, drawn round so that text can wrap tightly to it. */
export function sunflower(): Uint8Array {
	const size = 440;
	const canvas = new Canvas(size, size, [255, 255, 255]);
	const centre = size / 2;
	for (let ring = 0; ring < 2; ring++) {
		const petals = 18;
		const reach = ring ? 150 : 176;
		const colour = ring ? tone('marigold') : mix(tone('marigold'), tone('terracotta'), 0.35);
		for (let index = 0; index < petals; index++) {
			const angle = ((index + ring * 0.5) / petals) * Math.PI * 2;
			const tip: [number, number] = [
				centre + Math.cos(angle) * reach,
				centre + Math.sin(angle) * reach,
			];
			const spread = 0.19;
			const base = (offset: number): [number, number] => [
				centre + Math.cos(angle + offset) * 96,
				centre + Math.sin(angle + offset) * 96,
			];
			canvas.polygon([base(-spread), tip, base(spread)], colour);
			canvas.circle(tip[0] * 0.75 + centre * 0.25, tip[1] * 0.75 + centre * 0.25, 20, colour);
		}
	}
	canvas.circle(centre, centre, 104, tone('bark'));
	// Seeds in a sunflower spiral.
	for (let seed = 1; seed < 150; seed++) {
		const radius = 7.6 * Math.sqrt(seed);
		const angle = seed * 2.39996;
		canvas.circle(
			centre + Math.cos(angle) * radius,
			centre + Math.sin(angle) * radius,
			3.4,
			hex('4A3524'),
		);
	}
	return canvas.toPng();
}
