import { describe, expect, it } from 'vitest';

import { archEllipsePath } from './text-warp-arch';

describe('archEllipsePath', () => {
	it('draws the upper half of the box ellipse, below the frame inset, at the default angle', () => {
		expect(archEllipsePath(200, 100, 0.5, undefined, 'up')).toBe(
			'M 0,56.3 A 100,50 0 0,1 200,56.3',
		);
		expect(archEllipsePath(200, 100, 0.5, 10800000, 'up')).toBe(
			archEllipsePath(200, 100, 0.5, undefined, 'up'),
		);
	});

	it('draws the lower half the other way round for arch down', () => {
		expect(archEllipsePath(200, 100, 0.5, undefined, 'down')).toBe(
			'M 0,58.5 A 100,50 0 0,0 200,58.5',
		);
	});

	it('puts following lines on smaller concentric ellipses', () => {
		const first = archEllipsePath(200, 100, 0, undefined, 'up')!;
		const last = archEllipsePath(200, 100, 1, undefined, 'up')!;
		const rx = (path: string) => Number(/A ([\d.]+),/u.exec(path)![1]);
		expect(rx(first)).toBeGreaterThan(rx(last));
	});

	it('leaves other start angles to the previous approximation', () => {
		expect(archEllipsePath(200, 100, 0.5, 5400000, 'up')).toBeUndefined();
	});
});
