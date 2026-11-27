/** The report's pictures, drawn pixel by pixel: the cover, the site plan and the flow map. */
import { Canvas, hex, mix, type Rgb } from '../../lib/art';
import { PALETTE } from './data';

const color = (name: keyof typeof PALETTE): Rgb => hex(PALETTE[name]);

/** The cover: layered swells under a low amber disc, in flat tones that compress well. */
export function coverArt(): Uint8Array {
	const width = 1240;
	const height = 700;
	const sky = color('deep');
	const canvas = new Canvas(width, height, sky);
	canvas.fill((_, v) => mix(hex('0A3A47'), color('deep'), Math.round(v * 6) / 6));
	// The disc and the rings of its pull on the water.
	const cx = width * 0.74;
	const cy = height * 0.36;
	for (let ring = 5; ring >= 1; ring--)
		canvas.ring(cx, cy, 70 + ring * 46, 2, color('sea'), 0.1 + 0.03 * (5 - ring));
	canvas.circle(cx, cy, 70, color('amber'));
	const tones: Rgb[] = [
		mix(color('deep'), color('teal'), 0.35),
		mix(color('deep'), color('teal'), 0.7),
		color('teal'),
		mix(color('teal'), color('sea'), 0.55),
		color('sea'),
		mix(color('sea'), color('mist'), 0.6),
	];
	tones.forEach((tone, layer) => {
		const base = height * (0.5 + layer * 0.085);
		const amplitude = 34 - layer * 3.5;
		const crest = (x: number) =>
			base +
			amplitude * Math.sin(x / (150 + layer * 22) + layer * 1.7) +
			amplitude * 0.45 * Math.sin(x / (61 + layer * 9) + layer * 0.6);
		// Analytic coverage: a polygon of this many points would be tested for every pixel.
		canvas.paint(
			{ x0: 0, y0: base - 2 * amplitude, x1: width, y1: height },
			tone,
			(x, y) => y - crest(x) + 0.5,
		);
		// A lighter line just under each crest reads as the lip of the swell.
		canvas.paint(
			{ x0: 0, y0: base - 2 * amplitude, x1: width, y1: base + 2 * amplitude + 8 },
			mix(tone, [255, 255, 255], 0.35),
			(x, y) => 0.5 * (1.3 - Math.abs(y - crest(x) - 5)),
		);
	});
	return canvas.toPng();
}

/** The shores of the Narrows as functions of x (0 to 1): the channel is the water between them. */
const northShore = (u: number) =>
	0.2 + 0.16 * Math.exp(-(((u - 0.52) / 0.2) ** 2)) + 0.03 * Math.sin(u * 9);
const southShore = (u: number) =>
	0.84 - 0.2 * Math.exp(-(((u - 0.5) / 0.24) ** 2)) + 0.025 * Math.sin(u * 7 + 1);

/** Turbine positions in the throat of the channel, as fractions of the picture. */
const TURBINES: readonly (readonly [number, number])[] = [
	[0.47, 0.44],
	[0.5, 0.53],
	[0.53, 0.44],
	[0.56, 0.53],
];

function shores(canvas: Canvas): void {
	const { width, height } = canvas;
	const sand = hex('EDE6D3');
	const edge = hex('C9BD9C');
	const all = { x0: 0, y0: 0, x1: width, y1: height };
	const north = (x: number) => northShore(x / width) * height;
	const south = (x: number) => southShore(x / width) * height;
	canvas.paint(all, sand, (x, y) => Math.max(north(x) - y, y - south(x)) + 0.5);
	canvas.paint(all, edge, (x, y) => 1.7 - Math.min(Math.abs(y - north(x)), Math.abs(y - south(x))));
}

function turbines(canvas: Canvas): void {
	for (const [u, v] of TURBINES) {
		const x = u * canvas.width;
		const y = v * canvas.height;
		canvas.circle(x, y, 11, [255, 255, 255]);
		canvas.circle(x, y, 8, color('amber'));
		canvas.circle(x, y, 2.5, color('ink'));
	}
}

/** Figure 1: the channel in plan, with depth bands, the array, the export cable and the profilers. */
export function sitePlan(): Uint8Array {
	const width = 1000;
	const height = 560;
	const canvas = new Canvas(width, height, hex('D5E9ED'));
	const bands: Rgb[] = [hex('D5E9ED'), hex('B4D7DF'), hex('8FC3CF'), color('sea'), hex('3F8FA2')];
	canvas.fill((u, v) => {
		const top = northShore(u);
		const bottom = southShore(u);
		// Depth grows from each shore to the middle of the channel.
		const across = (v - top) / (bottom - top);
		const depth = Math.max(0, Math.min(1, 1 - Math.abs(across * 2 - 1) ** 1.6));
		return bands[Math.min(bands.length - 1, Math.floor(depth * bands.length))]!;
	});
	shores(canvas);
	// The export cable runs from the array to the substation on the north quay.
	const quay: [number, number] = [0.62 * width, northShore(0.62) * height - 20];
	const hub: [number, number] = [0.515 * width, 0.485 * height];
	canvas.line(hub[0], hub[1], 0.6 * width, 0.42 * height, 3, color('rust'));
	canvas.line(0.6 * width, 0.42 * height, quay[0], quay[1], 3, color('rust'));
	for (const [u, v] of TURBINES)
		canvas.line(u * width, v * height, hub[0], hub[1], 2, color('rust'), 0.8);
	canvas.rect(quay[0] - 13, quay[1] - 13, 26, 26, color('ink'), 1, 3);
	canvas.rect(quay[0] - 7, quay[1] - 7, 14, 14, color('amber'), 1, 2);
	turbines(canvas);
	// Two seabed current profilers, one each side of the array.
	for (const u of [0.34, 0.69]) {
		const x = u * width;
		const y = ((northShore(u) + southShore(u)) / 2) * height;
		const diamond = (size: number): [number, number][] => [
			[x, y - size],
			[x + size, y],
			[x, y + size],
			[x - size, y],
		];
		canvas.polygon(diamond(12), color('ink'));
		canvas.polygon(diamond(8), [255, 255, 255]);
	}
	// Flood-direction chevrons along the channel axis.
	for (const u of [0.1, 0.18, 0.26, 0.78, 0.86, 0.94]) {
		const x = u * width;
		const y = ((northShore(u) + southShore(u)) / 2) * height;
		canvas.line(x - 9, y - 9, x, y, 3, [255, 255, 255], 0.85);
		canvas.line(x - 9, y + 9, x, y, 3, [255, 255, 255], 0.85);
	}
	// North arrow and a 500 m scale bar in alternating blocks.
	canvas.polygon(
		[
			[940, 34],
			[954, 74],
			[940, 64],
			[926, 74],
		],
		color('ink'),
	);
	for (let block = 0; block < 5; block++)
		canvas.rect(40 + block * 40, 520, 40, 8, block % 2 ? [255, 255, 255] : color('ink'));
	return canvas.toPng();
}

/** Figure 4: modelled flow speed at peak spring flood, in seven bands of 0.5 m/s. */
export function flowMap(): Uint8Array {
	const width = 1000;
	const height = 560;
	const canvas = new Canvas(width, height, hex('E8F1F3'));
	const ramp: Rgb[] = [
		hex('E8F1F3'),
		hex('C4DEE4'),
		hex('9AC8D2'),
		color('sea'),
		color('teal'),
		color('deep'),
		color('amber'),
	];
	const throat = southShore(0.5) - northShore(0.52);
	canvas.fill((u, v) => {
		const top = northShore(u);
		const bottom = southShore(u);
		const across = (v - top) / (bottom - top);
		if (across < 0 || across > 1) return ramp[0]!;
		// Continuity speeds the flow up where the channel narrows; friction slows it at the shores.
		const speed = 3.47 * (throat / (bottom - top)) ** 0.9 * (1 - Math.abs(across * 2 - 1) ** 2.4);
		return ramp[Math.max(0, Math.min(ramp.length - 1, Math.floor(speed / 0.5)))]!;
	});
	shores(canvas);
	turbines(canvas);
	// The legend: the seven bands from slow to fast.
	ramp.forEach((tone, index) => {
		canvas.rect(40 + index * 34, 512, 34, 16, tone);
	});
	canvas.rect(39, 511, 34 * ramp.length + 2, 1.5, color('ink'));
	canvas.rect(39, 527.5, 34 * ramp.length + 2, 1.5, color('ink'));
	canvas.rect(39, 511, 1.5, 18, color('ink'));
	canvas.rect(39 + 34 * ramp.length + 0.5, 511, 1.5, 18, color('ink'));
	return canvas.toPng();
}
