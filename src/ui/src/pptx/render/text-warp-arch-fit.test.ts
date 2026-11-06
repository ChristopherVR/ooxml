import { describe, expect, it } from 'vitest';

import {
	archFitSweep,
	circleFitSweep,
	warpFitCapacity,
	warpFitScale,
	warpParagraphLayouts,
} from './text-warp-arch-fit';

describe('arc fit sweep', () => {
	it('leaves the whole half ellipse until the start angle nears the flip', () => {
		for (const degrees of [90, 120, 180, 240]) {
			expect(archFitSweep(degrees * 60000, 'up')).toBe(degrees <= 180 ? 180 : 540 - 2 * degrees);
		}
		expect(archFitSweep(undefined, 'up')).toBe(180);
		expect(archFitSweep(undefined, 'down')).toBe(180);
	});

	it('shrinks to nothing at 270 and 90 degrees, the other half turned the same way', () => {
		expect(archFitSweep(255 * 60000, 'up')).toBe(30);
		expect(archFitSweep(75 * 60000, 'up')).toBe(30);
		expect(archFitSweep(60 * 60000, 'down')).toBe(60);
		expect(archFitSweep(270 * 60000, 'up')).toBe(180);
		expect(archFitSweep(300 * 60000, 'up')).toBe(180);
	});

	it('shrinks Circle toward 180 and 360 degrees', () => {
		expect(circleFitSweep(undefined)).toBe(360);
		expect(circleFitSweep(330 * 60000)).toBe(60);
		expect(circleFitSweep(150 * 60000)).toBe(60);
		expect(circleFitSweep(90 * 60000)).toBe(180);
	});
});

describe('warpFitCapacity', () => {
	it('is the length of the half ellipse by default and falls with the sweep', () => {
		// Box 391 px by 142 px: half of the ellipse perimeter is about 442 px.
		const whole = warpFitCapacity('textArchUp', 391, 142, 0, undefined)!;
		expect(whole).toBeGreaterThan(436);
		expect(whole).toBeLessThan(448);
		// PowerPoint left 102 px of room at 75 degrees: 30 degrees about the bottom.
		const part = warpFitCapacity('textArchUp', 391, 142, 0, 255 * 60000)!;
		expect(part).toBeGreaterThan(95);
		expect(part).toBeLessThan(110);
	});

	it('has no capacity for a preset that does not fit its text to the path', () => {
		expect(warpFitCapacity('textWave1', 391, 142, 0, undefined)).toBeUndefined();
	});
});

describe('warpFitScale', () => {
	it('only ever shrinks', () => {
		expect(warpFitScale(400, 300)).toBe(1);
		expect(warpFitScale(300, 400)).toBe(0.75);
		expect(warpFitScale(undefined, 400)).toBe(1);
		expect(warpFitScale(300, undefined)).toBe(1);
	});
});

describe('warpParagraphLayouts', () => {
	const paragraph = (text: string) => ({ segments: [{ text, style: { fontSize: 20 } }] });

	it('gives every line its own path, one pitch (1.2 em) apart', () => {
		const layouts = warpParagraphLayouts(
			'textArchUp',
			200,
			100,
			[paragraph('a'), paragraph('b')],
			undefined,
			undefined,
			{ fontSize: 20 },
		);
		const rx = (d: string) => Number(/A ([\d.]+),/u.exec(d)![1]);
		expect(rx(layouts[0]!.d) - rx(layouts[1]!.d)).toBeCloseTo(24, 5);
	});
});
