import { describe, expect, it } from 'vitest';

import { easedProgress, easingSegment, timeAtProgress, type BezierPoints } from './morph-easing';
import {
	MORPH_CROSSFADE_EASING,
	MORPH_EASING,
	MORPH_EASING_POINTS,
	MORPH_FADE_IN_EASING,
	MORPH_FADE_IN_START_PERCENT,
	MORPH_FADE_IN_START_PROGRESS,
	MORPH_FADE_OUT_EASING,
	MORPH_FADE_OUT_END_PERCENT,
	MORPH_FADE_OUT_END_PROGRESS,
	MORPH_FADE_OUT_HOLD_PERCENT,
} from './morph-types';

function parse(easing: string): BezierPoints {
	const match = /^cubic-bezier\(([^)]*)\)$/u.exec(easing);
	expect(match).not.toBeNull();
	const values = (match?.[1] ?? '').split(',').map(Number);
	expect(values).toHaveLength(4);
	return values as unknown as BezierPoints;
}

describe('morph easing arithmetic', () => {
	it('evaluates a cubic-bezier at its ends and on the diagonal', () => {
		expect(easedProgress(0, 0.4, 0, 0.2, 1)).toBe(0);
		expect(easedProgress(1, 0.4, 0, 0.2, 1)).toBe(1);
		expect(easedProgress(0.3, 1 / 3, 1 / 3, 2 / 3, 2 / 3)).toBeCloseTo(0.3, 6);
		// CSS `ease` at half time, a published value.
		expect(easedProgress(0.5, 0.25, 0.1, 0.25, 1)).toBeCloseTo(0.8024, 3);
	});

	it('inverts the curve', () => {
		for (const progress of [0.05, 1 / 3, 0.5, 2 / 3, 0.95]) {
			const time = timeAtProgress(progress, MORPH_EASING_POINTS);
			expect(easedProgress(time, ...MORPH_EASING_POINTS)).toBeCloseTo(progress, 6);
		}
	});

	it('slices the curve so a keyframe interval stays linear in its progress', () => {
		const segment = easingSegment(MORPH_EASING_POINTS, 0.2, 0.7);
		const slice = parse(segment.easing);
		for (let i = 1; i < 10; i++) {
			const time = segment.startTime + ((segment.endTime - segment.startTime) * i) / 10;
			const whole = easedProgress(time, ...MORPH_EASING_POINTS);
			const local = easedProgress(i / 10, ...slice);
			expect(0.2 + 0.5 * local).toBeCloseTo(whole, 3);
		}
	});
});

describe('PowerPoint 16 morph timing (issue #141, measured)', () => {
	// Normalised progress of a matched rectangle's travel, size and rotation
	// read off PowerPoint's own CreateVideo export at 60fps, identical at 1s and
	// 2s and for an 840pt and an 80pt move (see MORPH_EASING_POINTS).
	const POWERPOINT_PROGRESS: ReadonlyArray<readonly [number, number]> = [
		[0.05, 0.021],
		[0.1, 0.075],
		[0.2, 0.236],
		[0.3, 0.419],
		[0.4, 0.588],
		[0.5, 0.726],
		[0.6, 0.832],
		[0.7, 0.911],
		[0.8, 0.96],
		[0.9, 0.988],
	];

	it('travels on the measured curve', () => {
		expect(MORPH_EASING).toBe('cubic-bezier(0.21, 0.02, 0.39, 1)');
		for (const [time, measured] of POWERPOINT_PROGRESS) {
			expect(Math.abs(easedProgress(time, ...MORPH_EASING_POINTS) - measured)).toBeLessThan(0.005);
		}
	});

	it('dissolves a matched pair on the same curve it travels on', () => {
		expect(MORPH_CROSSFADE_EASING).toBe(MORPH_EASING);
	});

	it('dissolves an unmatched shape out over the first third of the progress', () => {
		expect(MORPH_FADE_OUT_HOLD_PERCENT).toBe(0);
		expect(MORPH_FADE_OUT_END_PROGRESS).toBeCloseTo(1 / 3, 9);
		expect(MORPH_FADE_OUT_END_PERCENT).toBeCloseTo(25.2, 0);
		const slice = parse(MORPH_FADE_OUT_EASING);
		// Measured alpha of a leaving rectangle (luma, 2s morph, 60fps).
		for (const [time, alpha] of [
			[0.075, 0.865],
			[0.125, 0.675],
			[0.175, 0.43],
			[0.225, 0.155],
		] as const) {
			const local = easedProgress(time / (MORPH_FADE_OUT_END_PERCENT / 100), ...slice);
			expect(Math.abs(1 - local - alpha)).toBeLessThan(0.02);
		}
	});

	it('dissolves an unmatched shape in over the last third of the progress', () => {
		expect(MORPH_FADE_IN_START_PROGRESS).toBeCloseTo(2 / 3, 9);
		expect(MORPH_FADE_IN_START_PERCENT).toBeCloseTo(45.5, 0);
		expect(MORPH_FADE_OUT_END_PERCENT).toBeLessThan(MORPH_FADE_IN_START_PERCENT);
		const slice = parse(MORPH_FADE_IN_EASING);
		const start = MORPH_FADE_IN_START_PERCENT / 100;
		for (const [time, alpha] of [
			[0.525, 0.26],
			[0.625, 0.565],
			[0.725, 0.775],
			[0.825, 0.91],
		] as const) {
			const local = easedProgress((time - start) / (1 - start), ...slice);
			expect(Math.abs(local - alpha)).toBeLessThan(0.02);
		}
	});
});
