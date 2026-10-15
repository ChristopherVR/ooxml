import { describe, expect, it } from 'vitest';
import { timingSafeEqual } from './timing-safe';

describe('timingSafeEqual', () => {
	it('compares content and length', () => {
		expect(timingSafeEqual(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2, 3]))).toBe(true);
		expect(timingSafeEqual(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2, 4]))).toBe(false);
		expect(timingSafeEqual(new Uint8Array([9, 2, 3]), new Uint8Array([1, 2, 3]))).toBe(false);
		expect(timingSafeEqual(new Uint8Array([1, 2]), new Uint8Array([1, 2, 3]))).toBe(false);
		expect(timingSafeEqual(new Uint8Array(0), new Uint8Array(0))).toBe(true);
	});
});
