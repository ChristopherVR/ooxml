/**
 * Case 01, Brindlemoor Bakehouse: the identity board (wheat mark, stamp, awning stripe, loaf
 * pictogram) and a packaging family. Flat colour throughout, so the PNGs stay small.
 */
import { Canvas, hex } from '../../lib/art';
import { ellipse, shadow, textLines, wheat } from './shapes';

const CREAM = hex('#F3E7CF');
const TERRACOTTA = hex('#B5532F');
const CRUST = hex('#7A4A2A');
const COCOA = hex('#3B2418');
const OCHRE = hex('#D6A24A');
const KRAFT = hex('#C9A77C');

/** A loaf seen from above: an oval with three diagonal scores. */
function loaf(canvas: Canvas, cx: number, cy: number, size: number): void {
	ellipse(canvas, cx, cy, size, size * 0.56, OCHRE, 1, -0.2);
	for (const offset of [-0.42, 0, 0.42])
		canvas.line(
			cx + offset * size - size * 0.1,
			cy + size * 0.2,
			cx + offset * size + size * 0.12,
			cy - size * 0.24,
			size * 0.07,
			COCOA,
		);
}

/** The round stamp: a ring, a wheat ear and two arcs of text bars. */
function stamp(canvas: Canvas, cx: number, cy: number, radius: number, ink = CRUST): void {
	canvas.ring(cx, cy, radius, radius * 0.045, ink);
	canvas.ring(cx, cy, radius * 0.86, radius * 0.018, ink);
	wheat(canvas, cx, cy + radius * 0.58, radius * 0.95, ink, radius / 190);
	for (let i = 0; i < 11; i++) {
		const angle = Math.PI * (1.12 + (0.76 * i) / 10);
		canvas.circle(
			cx + Math.cos(angle) * radius * 0.72,
			cy + Math.sin(angle) * radius * 0.72,
			radius * 0.022,
			ink,
		);
	}
}

export function bakehouseBoard(): Uint8Array {
	const canvas = new Canvas(960, 600, CREAM);
	// Left panel: the wheat mark, large, on terracotta.
	canvas.rect(0, 0, 372, 600, TERRACOTTA);
	wheat(canvas, 186, 548, 440, CREAM, 1.5);
	// Middle panel: the stamp and the wordmark, as bars.
	stamp(canvas, 536, 232, 128);
	canvas.rect(456, 420, 160, 14, COCOA, 1, 3);
	textLines(canvas, 456, 452, [116, 148, 92], 5, 15, CRUST, 0.75);
	canvas.rect(456, 528, 44, 4, TERRACOTTA);
	// Right panel: the awning stripe above, the loaf pictogram below.
	canvas.rect(700, 0, 260, 300, OCHRE);
	for (let i = 0; i < 7; i++) if (i % 2 === 0) canvas.rect(700 + i * 37.2, 0, 37.2, 300, CREAM);
	canvas.rect(700, 300, 260, 300, COCOA);
	loaf(canvas, 830, 432, 74);
	textLines(canvas, 782, 520, [96, 64], 4, 13, CREAM, 0.7);
	return canvas.toPng();
}

export function bakehousePackaging(): Uint8Array {
	const canvas = new Canvas(640, 400, hex('#E6D8C4'));
	canvas.rect(0, 286, 640, 114, hex('#DBCAB2'));
	// The paper bag with a folded top.
	shadow(canvas, 168, 338, 104, 16, 0.3);
	canvas.polygon(
		[
			[92, 104],
			[232, 104],
			[244, 336],
			[80, 336],
		],
		KRAFT,
	);
	canvas.polygon(
		[
			[92, 104],
			[232, 104],
			[226, 76],
			[98, 76],
		],
		hex('#B8956B'),
	);
	canvas.rect(84, 196, 156, 74, CREAM);
	wheat(canvas, 162, 262, 52, CRUST, 0.2);
	// The pastry box with a band and the stamp.
	shadow(canvas, 380, 340, 112, 15, 0.3);
	canvas.rect(288, 190, 184, 148, TERRACOTTA, 1, 4);
	canvas.rect(288, 190, 184, 34, hex('#9C4426'), 1, 4);
	canvas.rect(288, 248, 184, 22, CREAM);
	stamp(canvas, 380, 296, 26, CREAM);
	// The coffee cup and a round label.
	shadow(canvas, 548, 342, 56, 11, 0.3);
	canvas.polygon(
		[
			[506, 224],
			[590, 224],
			[578, 338],
			[518, 338],
		],
		CREAM,
	);
	canvas.rect(500, 212, 96, 16, COCOA, 1, 5);
	canvas.rect(512, 262, 72, 34, OCHRE);
	canvas.circle(548, 122, 44, COCOA);
	stamp(canvas, 548, 122, 36, OCHRE);
	return canvas.toPng();
}
