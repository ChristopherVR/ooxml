import { describe, expect, it } from 'vitest';
import { recognizeInkShape } from '../ink-recognition';
import { DEFAULT_INK_PENS, normalizeInkPens } from '../ink-pens';
import {
	eraseInkStrokes,
	inkPathPoints,
	inkStrokeFromSamples,
	inkStrokesBounds,
	lassoInkStrokes,
	transformInkStrokes,
} from './edit';
import { inkShapeOf } from './to-shape';
import type { InkStroke } from './types';

const pen = { kind: 'pen' as const, color: '#000000', width: 2 };
const line = (x1: number, y1: number, x2: number, y2: number): InkStroke =>
	inkStrokeFromSamples(
		[
			{ x: x1, y: y1 },
			{ x: x2, y: y2 },
		],
		pen,
	)!;

describe('ink edits', () => {
	it('makes a stroke from samples, dropping repeats and keeping varying pressure', () => {
		const stroke = inkStrokeFromSamples(
			[
				{ x: 0, y: 0, pressure: 0.2 },
				{ x: 0.1, y: 0.1, pressure: 0.3 },
				{ x: 10, y: 0, pressure: 0.9 },
			],
			{ kind: 'highlighter', color: '#FFFF00', width: 12 },
		)!;
		expect(stroke.path).toBe('M 0 0 L 10 0');
		expect(stroke.pressures).toEqual([0.2, 0.9]);
		expect(stroke.opacity).toBe(0.4);
		const mouse = inkStrokeFromSamples(
			[
				{ x: 1, y: 1 },
				{ x: 4, y: 4 },
			],
			pen,
		)!;
		expect(mouse.pressures).toBeUndefined();
		expect(inkStrokeFromSamples([], pen)).toBeUndefined();
		// A tap is a dot.
		expect(inkPathPoints(inkStrokeFromSamples([{ x: 3, y: 3 }], pen)!.path)).toHaveLength(2);
	});

	it('measures, moves and scales strokes', () => {
		const strokes = [line(10, 10, 30, 20)];
		expect(inkStrokesBounds(strokes)).toEqual({ x: 9, y: 9, width: 22, height: 12 });
		const moved = transformInkStrokes(strokes, -10, -10, 2, 1);
		expect(moved[0]?.path).toBe('M 0 0 L 40 10');
	});

	it('erases whole strokes the eraser touches', () => {
		const strokes = [line(0, 0, 100, 0), line(0, 50, 100, 50)];
		expect(eraseInkStrokes(strokes, { x: 50, y: 3 }, 2)).toEqual([strokes[1]]);
		expect(eraseInkStrokes(strokes, { x: 50, y: 25 }, 2)).toHaveLength(2);
	});

	it('lassos the strokes inside the outline', () => {
		const strokes = [line(10, 10, 20, 20), line(100, 100, 120, 120)];
		const outline = [
			{ x: 0, y: 0 },
			{ x: 50, y: 0 },
			{ x: 50, y: 50 },
			{ x: 0, y: 50 },
		];
		expect(lassoInkStrokes(strokes, outline)).toEqual([0]);
	});

	it('turns a recognised line into an unfilled line with an arrowhead for an arrow', () => {
		const recognized = recognizeInkShape([inkPathPoints(line(0, 40, 100, 0).path)])!;
		const shape = inkShapeOf(recognized, '#E81123', 3);
		expect(shape.preset).toBe('line');
		expect(shape.flipV).toBe(true);
		expect(shape.line.widthEmu).toBe(28575);
		expect(shape.line.fill).toEqual({
			kind: 'solid',
			color: { kind: 'srgb', value: 'E81123', transforms: [] },
		});
	});
});

describe('ink pens', () => {
	it('keeps valid pens of a stored gallery', () => {
		expect(DEFAULT_INK_PENS.map((p) => p.kind)).toEqual([
			'pen',
			'pen',
			'pen',
			'pencil',
			'highlighter',
		]);
		expect(
			normalizeInkPens([
				{ kind: 'pen', color: '#123456', width: 3 },
				{ kind: 'marker', color: '#123456', width: 3 },
				{ kind: 'pen', color: 'red', width: 3 },
				{ kind: 'highlighter', color: '#FFFF00', width: 22 },
			]),
		).toEqual([{ kind: 'pen', color: '#123456', width: 3 }]);
		expect(
			normalizeInkPens([{ kind: 'highlighter', color: '#FFFF00', width: 22 }], 32),
		).toHaveLength(1);
	});
});
