/**
 * The studio's own pictures: the cover composition (a thimble, a needle's thread and three
 * fields of the palette) and the abstract portraits of the team.
 */
import { Canvas, hex, mix, type Rgb } from '../../lib/art';
import { curve, ellipse } from './shapes';
import { C } from './style';

const BONE = hex(C.bone);
const INK = hex(C.ink);

export function coverComposition(): Uint8Array {
	const canvas = new Canvas(720, 828, hex(C.clay));
	// A sage arch, an ink-blue quarter disc and the vermilion sun between them.
	canvas.circle(250, 300, 170, hex(C.sage));
	canvas.rect(80, 300, 340, 528, hex(C.sage));
	canvas.circle(720, 828, 400, hex(C.blue));
	canvas.circle(520, 236, 104, hex(C.vermilion));
	// The thimble: a tapered cup with a rim and rows of dimples.
	const cup: [number, number][] = [
		[168, 730],
		[332, 730],
		[310, 560],
		[190, 560],
	];
	canvas.polygon(cup, INK);
	ellipse(canvas, 250, 560, 60, 44, INK);
	canvas.rect(156, 716, 188, 26, INK, 1, 6);
	for (let row = 0; row < 5; row++)
		for (let dot = 0; dot < 5 - (row % 2); dot++) {
			const y = 578 + row * 26;
			const reach = 44 + row * 4;
			const x = 250 + (dot - (4 - (row % 2)) / 2) * (reach / 2.2);
			canvas.circle(x, y, 4.2, mix(INK, BONE, 0.32));
		}
	// The thread: one line that loops across the fields and ends in a needle's eye.
	const thread: [number, number][] = [];
	for (let i = 0; i <= 120; i++) {
		const t = i / 120;
		thread.push([
			40 + t * 600 + Math.sin(t * Math.PI * 3) * 46,
			120 + t * 420 + Math.sin(t * Math.PI * 2 + 0.6) * 150,
		]);
	}
	curve(canvas, thread, 3, BONE);
	const end = thread[thread.length - 1]!;
	canvas.ring(end[0] + 12, end[1] + 10, 13, 3.4, BONE);
	return canvas.toPng();
}

interface Sitter {
	ground: string;
	skin: string;
	hair: string;
	top: string;
	/** 0 cropped, 1 bob, 2 bun, 3 curls, 4 long. */
	style: number;
}

const SITTERS: Sitter[] = [
	{ ground: C.clay, skin: '#E0B08C', hair: '#3A2A22', top: C.vermilion, style: 2 },
	{ ground: C.sage, skin: '#7A4E36', hair: '#17110E', top: C.bone, style: 0 },
	{ ground: '#C7CDDD', skin: '#F0CDAE', hair: '#C9A45C', top: C.blue, style: 4 },
	{ ground: '#E2CE9F', skin: '#D3A27C', hair: '#1D1D1B', top: C.moss, style: 1 },
	{ ground: '#DDB9A8', skin: '#5F3D2B', hair: '#241812', top: C.wheat, style: 3 },
];

function hair(canvas: Canvas, style: number, color: Rgb): void {
	switch (style) {
		case 1:
			canvas.rect(88, 96, 124, 132, color, 1, 56);
			break;
		case 2:
			canvas.circle(150, 62, 30, color);
			ellipse(canvas, 150, 128, 58, 60, color);
			break;
		case 3:
			for (const [x, y, r] of [
				[104, 112, 30],
				[132, 86, 32],
				[170, 86, 32],
				[198, 112, 30],
				[100, 148, 24],
				[202, 148, 24],
			] as const)
				canvas.circle(x, y, r, color);
			break;
		case 4:
			canvas.rect(86, 92, 128, 196, color, 1, 60);
			break;
		default:
			ellipse(canvas, 150, 122, 56, 54, color);
	}
}

/** An abstract bust: flat shapes for shoulders, neck, head and hair, no features. */
export function portrait(index: number): Uint8Array {
	const sitter = SITTERS[index % SITTERS.length]!;
	const canvas = new Canvas(300, 360, hex(sitter.ground));
	const skin = hex(sitter.skin);
	hair(canvas, sitter.style, hex(sitter.hair));
	ellipse(canvas, 150, 420, 150, 150, hex(sitter.top));
	canvas.rect(128, 190, 44, 70, mix(skin, [0, 0, 0], 0.12));
	ellipse(canvas, 150, 150, 46, 56, skin);
	// A fringe over the forehead for the styles that have one.
	if (sitter.style === 1 || sitter.style === 4)
		canvas.rect(104, 94, 92, 30, hex(sitter.hair), 1, 15);
	if (sitter.style === 0) ellipse(canvas, 150, 104, 47, 20, hex(sitter.hair));
	return canvas.toPng();
}
