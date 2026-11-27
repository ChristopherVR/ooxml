import { describe, expect, it } from 'vitest';
import {
	contentControlTagMarker,
	contentControlTagOutline,
	contentControlTagOutlineWidth,
	contentControlTagTitleOffset,
	tagPathData,
	type TagPathCommand,
} from './content-control-tag-shape';

/** Points along the outline, sampling each curve. */
function sample(commands: readonly TagPathCommand[]): [number, number][] {
	const points: [number, number][] = [];
	let at: [number, number] = [0, 0];
	for (const command of commands) {
		if (command[0] !== 'C') {
			at = [command[1], command[2]];
			points.push(at);
			continue;
		}
		const [, x1, y1, x2, y2, x, y] = command;
		for (let step = 1; step <= 200; step++) {
			const t = step / 200;
			const u = 1 - t;
			points.push([
				u * u * u * at[0] + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x,
				u * u * u * at[1] + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y,
			]);
		}
		at = [x, y];
	}
	return points;
}

function distance(points: [number, number][], [x, y]: readonly [number, number]): number {
	return Math.min(...points.map(([px, py]) => Math.hypot(px - x, py - y)));
}

// Vertices of Word's 600 dpi polylines in inline-11-tags.pdf, from the tag's top left, y down.
const wordStart: [number, number][] = [
	[2.28, 13.44],
	[1.56, 12.48],
	[0.96, 11.28],
	[0.48, 9.84],
	[0.12, 8.4],
	[0, 6.72],
	[0.24, 4.32],
	[0.72, 2.76],
	[1.32, 1.44],
	[2.28, 0],
	[18.48, 0],
	[16.92, 2.88],
	[16.2, 6.72],
	[16.56, 9.84],
	[18.48, 13.44],
];
const wordInlineEnd: [number, number][] = [
	[0, 13.44],
	[0.96, 12],
	[2.28, 6.72],
	[1.56, 10.56],
	[0, 0],
	[7.56, 0],
	[9.84, 6.72],
	[9.12, 10.56],
	[7.56, 13.44],
];

describe('content control tag outline', () => {
	it("follows Word's opening tag polyline with both edges bowing left", () => {
		const outline = contentControlTagOutline('start', 18.48);
		const points = sample(outline);
		for (const point of wordStart) expect(distance(points, point)).toBeLessThan(0.13);
		expect(Math.min(...points.map(([x]) => x))).toBeCloseTo(0, 2);
		expect(Math.max(...points.map(([x]) => x))).toBeCloseTo(18.48, 2);
	});

	it("follows Word's unlabelled closing tag, bowing right inside its advance", () => {
		const width = contentControlTagOutlineWidth('end', 11.5, false);
		expect(width).toBe(9.84);
		expect(contentControlTagOutlineWidth('end', 56.64, true)).toBe(56.64);
		const points = sample(contentControlTagOutline('end', width));
		for (const point of wordInlineEnd) expect(distance(points, point)).toBeLessThan(0.13);
	});

	it('writes SVG path data and the title and marker positions', () => {
		expect(tagPathData(contentControlTagOutline('start', 20)).startsWith('M 2.28 13.44 C')).toBe(
			true,
		);
		expect(tagPathData(contentControlTagOutline('end', 10)).endsWith(' Z')).toBe(true);
		expect([
			contentControlTagTitleOffset('start', false),
			contentControlTagTitleOffset('start', true),
			contentControlTagTitleOffset('end', true),
		]).toEqual([6.72, 9, 5.04]);
		// block-tags.pdf: start tag at 45, end tag 98.784 to 155.42, box tops 743.04 and 724.32.
		const start = contentControlTagMarker('start', 56.66);
		expect({
			arrow: start.arrow.map((point) => point.map((value) => Number(value.toFixed(3)))),
			bar: start.bar,
		}).toEqual({
			arrow: [
				[2.34, 6.78],
				[4.5, 5.1],
				[4.5, 8.46],
			],
			bar: { x: 5.58, y: 5.1, width: 0.6, height: 3.36 },
		});
		const end = contentControlTagMarker('end', 56.64);
		expect(end.bar.x).toBeCloseTo(148.76 - 98.784, 2);
		expect(at(end.arrow, 0)[0]).toBeCloseTo(152.6 - 98.784, 2);
		expect(at(end.arrow, 1)[0]).toBeCloseTo(150.44 - 98.784, 2);
	});
});

function at<T>(items: readonly T[], index: number): T {
	const item = items[index];
	if (item === undefined) throw new Error(`Missing item ${index}`);
	return item;
}
