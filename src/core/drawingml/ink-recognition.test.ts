import { describe, expect, it } from 'vitest';

import { recognizeInkShape } from './ink-recognition';
import { slideInkReplayTimeline } from './ink-replay';

const points = (...coordinates: [number, number][]) => coordinates.map(([x, y]) => ({ x, y }));

describe('ink shape recognition', () => {
	it('classifies a rectangle, ellipse and triangle as DrawingML presets', () => {
		const rect = points([0, 0], [80, 1], [80, 40], [0, 40], [1, 1]);
		const ellipse = Array.from({ length: 41 }, (_, i) => ({
			x: 40 * Math.cos((i * Math.PI) / 20),
			y: 25 * Math.sin((i * Math.PI) / 20),
		}));
		const triangle = points([40, 0], [80, 48], [0, 48], [40, 0]);
		expect(recognizeInkShape([rect])?.kind).toBe('rect');
		expect(recognizeInkShape([ellipse])?.kind).toBe('ellipse');
		expect(recognizeInkShape([triangle])).toMatchObject({ kind: 'triangle', rotation: 0 });
	});

	it('keeps line and arrow direction and rejects an open corner', () => {
		const shaft = points([90, 30], [20, 30]);
		const head = points([40, 15], [20, 30], [40, 45]);
		expect(recognizeInkShape([shaft])).toStrictEqual({
			kind: 'line',
			start: shaft[0],
			end: shaft[1],
		});
		expect(recognizeInkShape([head, shaft])).toStrictEqual({
			kind: 'arrow',
			start: shaft[0],
			end: shaft[1],
		});
		expect(recognizeInkShape([points([0, 0], [80, 0], [80, 40])])).toBeUndefined();
	});

	it('bounds the samples it will analyze', () => {
		expect(
			recognizeInkShape([Array.from({ length: 4097 }, (_, x) => ({ x, y: 0 }))]),
		).toBeUndefined();
	});
});

describe('slide ink timing', () => {
	it('orders timestamped strokes across parts and retains the recorded gap', () => {
		expect(
			slideInkReplayTimeline([
				{ pointTimestamps: [1600, 1700] },
				{ pointTimestamps: [1000, 1200] },
			]),
		).toStrictEqual([
			{ strokeIndex: 1, startOffsetMs: 0, durationMs: 200 },
			{ strokeIndex: 0, startOffsetMs: 600, durationMs: 100 },
		]);
	});

	it('cascades traces with equal context times and retains a duration beside untimed ink', () => {
		expect(
			slideInkReplayTimeline([
				{ startTimestamp: 300, durationMs: 100 },
				{ startTimestamp: 300 },
				{},
			]),
		).toStrictEqual([
			{ strokeIndex: 0, startOffsetMs: 0, durationMs: 100 },
			{ strokeIndex: 1, startOffsetMs: 300, durationMs: 600 },
			{ strokeIndex: 2, startOffsetMs: 1100, durationMs: 600 },
		]);
	});
});
