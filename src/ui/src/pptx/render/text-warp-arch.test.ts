import { describe, expect, it } from 'vitest';

import {
	archEllipsePath,
	archOnUpperHalf,
	circleEllipsePath,
	normalizeWarpAngle,
	warpLineOutward,
} from './text-warp-arch';

describe('archEllipsePath', () => {
	it('draws the upper half of the box ellipse, below the frame inset, at the default angle', () => {
		expect(archEllipsePath(200, 100, 0, undefined, 'up')).toBe('M 0,56.3 A 100,50 0 0,1 200,56.3');
		expect(archEllipsePath(200, 100, 0, 10800000, 'up')).toBe(
			archEllipsePath(200, 100, 0, undefined, 'up'),
		);
	});

	it('draws the lower half the other way round for arch down', () => {
		expect(archEllipsePath(200, 100, 0, undefined, 'down')).toBe(
			'M 0,58.5 A 100,50 0 0,0 200,58.5',
		);
	});

	it('puts lines further out on larger concentric ellipses, one pitch apart', () => {
		const rx = (path: string) => Number(/A ([\d.]+),/u.exec(path)![1]);
		expect(rx(archEllipsePath(200, 100, 0, undefined, 'up'))).toBe(100);
		expect(rx(archEllipsePath(200, 100, 29, undefined, 'up'))).toBe(129);
	});

	it('stacks the lines away from the box centre: the base path carries the last line (the first for arch down)', () => {
		expect(warpLineOutward('textArchUp', 2, 3, 29)).toBe(0);
		expect(warpLineOutward('textArchUp', 0, 3, 29)).toBe(58);
		expect(warpLineOutward('textButton', 0, 2, 29)).toBe(29);
		expect(warpLineOutward('textCircle', 0, 2, 29)).toBe(29);
		expect(warpLineOutward('textArchDown', 0, 3, 29)).toBe(0);
		expect(warpLineOutward('textArchDown', 2, 3, 29)).toBe(58);
		expect(warpLineOutward('textArchUp', 0, 1, 29)).toBe(0);
	});

	it('picks the half of the ellipse from the start angle, not a rotation of the arc', () => {
		// Measured on PowerPoint 16: 90..269 degrees upper half, 270..89 lower half.
		for (const degrees of [90, 120, 180, 240, 269]) {
			expect(archOnUpperHalf(degrees * 60000, 'up')).toBe(true);
			expect(archEllipsePath(200, 100, 0, degrees * 60000, 'up')).toBe(
				archEllipsePath(200, 100, 0, undefined, 'up'),
			);
		}
		for (const degrees of [270, 300, 0, 30, 89]) {
			expect(archOnUpperHalf(degrees * 60000, 'up')).toBe(false);
		}
		expect(archOnUpperHalf(undefined, 'down')).toBe(false);
	});

	it('turns arch up half a turn about the box centre for a start angle on the lower half', () => {
		const up = archEllipsePath(200, 100, 0, undefined, 'up');
		expect(archEllipsePath(200, 100, 0, 16200000, 'up')).toBe('M 200,43.7 A 100,50 0 0,1 0,43.7');
		expect(archEllipsePath(200, 100, 0, 0, 'up')).not.toBe(up);
		expect(archEllipsePath(200, 100, 0, 5400000, 'up')).toBe(up);
	});

	it('turns arch down to the upper half for a start angle from 90 to 269 degrees', () => {
		expect(archEllipsePath(200, 100, 0, 10800000, 'down')).toBe('M 200,41.5 A 100,50 0 0,0 0,41.5');
		expect(archEllipsePath(200, 100, 0, 0, 'down')).toBe(
			archEllipsePath(200, 100, 0, undefined, 'down'),
		);
	});

	it('normalises an angle into one turn', () => {
		expect(normalizeWarpAngle(-5400000)).toBe(270);
		expect(normalizeWarpAngle(21600000 + 5400000)).toBe(90);
	});
});

describe('circleEllipsePath', () => {
	const points = (path: string): number[][] =>
		[...path.matchAll(/(?:M|A [\d.]+,[\d.]+ 0 0,1) ([-\d.]+),([-\d.]+)/gu)].map((m) => [
			Number(m[1]),
			Number(m[2]),
		]);

	it('starts at the left tip and runs once round clockwise, so the centred text sits on the right', () => {
		const path = circleEllipsePath(200, 100, 0, undefined);
		const [start, opposite] = points(path) as [number[], number[]];
		expect(start[0]).toBeLessThan(10);
		expect(opposite[0]).toBeGreaterThan(190);
		expect(path).toContain('0 0,1');
		// The text is centred at half the path length: the second point of the loop.
		expect(path.split('A')).toHaveLength(3);
	});

	it('centres the text on the left tip for a start angle from 0 to 179 degrees', () => {
		const left = points(circleEllipsePath(200, 100, 0, undefined))[0]!;
		for (const degrees of [0, 45, 90, 179]) {
			const turned = points(circleEllipsePath(200, 100, 0, degrees * 60000))[0]!;
			expect(turned[0]).toBeGreaterThan(190);
			expect(left[0]).toBeLessThan(10);
		}
		for (const degrees of [180, 270, 359]) {
			expect(points(circleEllipsePath(200, 100, 0, degrees * 60000))[0]).toEqual(left);
		}
	});
});
