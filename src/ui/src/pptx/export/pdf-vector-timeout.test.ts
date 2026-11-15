import { afterEach, describe, expect, it, vi } from 'vitest';

import { settleWithin } from './pdf-vector-timeout';

describe('settleWithin', () => {
	afterEach(() => {
		vi.useRealTimers();
	});

	it('passes a value that settles in time through', async () => {
		await expect(settleWithin(Promise.resolve(3), 50, 'late')).resolves.toBe(3);
	});

	it('yields the fallback for a rejection', async () => {
		await expect(settleWithin(Promise.reject(new Error('no')), 50, 'failed')).resolves.toBe(
			'failed',
		);
	});

	it('yields the fallback when the work never settles', async () => {
		vi.useFakeTimers();
		const result = settleWithin(new Promise<number>(() => undefined), 1000, undefined);
		await vi.advanceTimersByTimeAsync(1000);
		await expect(result).resolves.toBeUndefined();
	});
});
