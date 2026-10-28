import { expect, it, vi } from 'vitest';
import { createRecoveryScheduler } from './recovery-scheduler';

it('ignores stale ticks after reconfiguration and disposal', () => {
	const callbacks: Array<() => void> = [];
	const tick = vi.fn();
	const scheduler = createRecoveryScheduler({
		tick,
		timers: {
			set: (callback) => {
				callbacks.push(callback);
				return callbacks.length;
			},
			clear: vi.fn(),
		},
	});
	scheduler.arm(100);
	scheduler.arm(200);
	callbacks[0]!();
	expect(tick).not.toHaveBeenCalled();
	callbacks[1]!();
	expect(tick).toHaveBeenCalledOnce();
	scheduler.dispose();
	callbacks[1]!();
	expect(tick).toHaveBeenCalledOnce();
});

it('serializes pending writes, reads later state, and skips queued writes after disposal', async () => {
	let finish!: () => void;
	const pending = new Promise<void>((resolve) => {
		finish = resolve;
	});
	const scheduler = createRecoveryScheduler({ tick: () => {} });
	const first = scheduler.run(() => pending);
	let revision = 1;
	const write = vi.fn(async () => revision);
	const second = scheduler.run(write);
	revision = 2;
	expect(write).not.toHaveBeenCalled();
	finish();
	await first;
	expect(await second).toBe(2);
	const next = scheduler.run(async () => {});
	const queued = scheduler.run(write);
	scheduler.dispose();
	await next;
	expect(await queued).toBeUndefined();
	expect(write).toHaveBeenCalledOnce();
});
