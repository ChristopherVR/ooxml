import { describe, expect, it } from 'vitest';

import { nudgeDelta } from './editor-geometry';

describe('nudgeDelta', () => {
	it('maps the four arrows to unit steps', () => {
		expect(nudgeDelta('ArrowLeft', false)).toStrictEqual({ dx: -1, dy: 0 });
		expect(nudgeDelta('ArrowRight', false)).toStrictEqual({ dx: 1, dy: 0 });
		expect(nudgeDelta('ArrowUp', false)).toStrictEqual({ dx: 0, dy: -1 });
		expect(nudgeDelta('ArrowDown', false)).toStrictEqual({ dx: 0, dy: 1 });
	});

	it('uses the 10px large step when shift is held', () => {
		expect(nudgeDelta('ArrowRight', true)).toStrictEqual({ dx: 10, dy: 0 });
	});

	it('returns null for non-arrow keys', () => {
		expect(nudgeDelta('Enter', false)).toBeNull();
	});
});
