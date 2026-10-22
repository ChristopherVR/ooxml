import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { flushWriteBackOnPageClose } from './write-back-lifecycle';
import { createWriteBackScheduler, writeBackStateOf, type WriteBackOptions } from './write-back';

function make(overrides: Partial<WriteBackOptions> = {}) {
	let n = 0;
	const statuses: string[] = [];
	const write = vi.fn((_bytes: Uint8Array): void | Promise<void> => {});
	const scheduler = createWriteBackScheduler({
		serialize: () => new Uint8Array([++n]),
		write,
		debounceMs: 100,
		onStatus: (s) => statuses.push(s.status),
		now: () => 42,
		...overrides,
	});
	return { scheduler, write, statuses };
}

beforeEach(() => {
	vi.useFakeTimers();
});
afterEach(() => {
	vi.useRealTimers();
});

describe('createWriteBackScheduler', () => {
	it('debounces with a trailing timer and treats a void write as acknowledged', async () => {
		const { scheduler, write, statuses } = make();
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(60);
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(60);
		expect(write).not.toHaveBeenCalled();
		expect(scheduler.state.status).toBe('pending');
		await vi.advanceTimersByTimeAsync(50);
		expect(write).toHaveBeenCalledTimes(1);
		expect(scheduler.state).toEqual({ status: 'saved', lastSavedAt: 42 });
		expect(statuses).toEqual(['pending', 'saving', 'saved']);
		expect(scheduler.dirty).toBe(false);
	});

	it('uses 5000 ms by default and falls back from invalid values', async () => {
		const { scheduler, write } = make({ debounceMs: Number.NaN });
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(4_999);
		expect(write).not.toHaveBeenCalled();
		await vi.advanceTimersByTimeAsync(2);
		expect(write).toHaveBeenCalledTimes(1);
	});

	it('awaits a returned promise as the acknowledgement', async () => {
		let release!: () => void;
		const { scheduler } = make({
			write: () =>
				new Promise<void>((resolve) => {
					release = resolve;
				}),
		});
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(scheduler.state.status).toBe('saving');
		release();
		await vi.advanceTimersByTimeAsync(0);
		expect(scheduler.state.status).toBe('saved');
	});

	it('keeps one write in flight and coalesces changes made during it into one follow-up', async () => {
		const releases: (() => void)[] = [];
		const { scheduler } = make({
			write: () => new Promise<void>((resolve) => releases.push(resolve)),
		});
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(releases).toHaveLength(1);
		scheduler.schedule();
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(10_000);
		expect(releases).toHaveLength(1);
		releases[0]!();
		await vi.advanceTimersByTimeAsync(0);
		expect(scheduler.state.status).toBe('pending');
		await vi.advanceTimersByTimeAsync(100);
		expect(releases).toHaveLength(2);
		releases[1]!();
		await vi.advanceTimersByTimeAsync(0);
		expect(scheduler.state.status).toBe('saved');
		await vi.advanceTimersByTimeAsync(10_000);
		expect(releases).toHaveLength(2);
	});

	it('retries with exponential backoff, then reports saved', async () => {
		let calls = 0;
		const { scheduler, statuses } = make({
			retry: { attempts: 4, baseDelayMs: 100, maxDelayMs: 150 },
			write: () => {
				calls++;
				return calls < 3 ? Promise.reject(new Error('offline')) : Promise.resolve();
			},
		});
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(calls).toBe(1);
		await vi.advanceTimersByTimeAsync(99);
		expect(calls).toBe(1);
		await vi.advanceTimersByTimeAsync(2);
		expect(calls).toBe(2);
		// The second wait is capped at maxDelayMs (150, not 200).
		await vi.advanceTimersByTimeAsync(151);
		expect(calls).toBe(3);
		expect(scheduler.state.status).toBe('saved');
		expect(statuses).not.toContain('failed');
	});

	it('reports failed with the error after the retries are spent, and recovers on the next change', async () => {
		let fail = true;
		const { scheduler } = make({
			retry: { attempts: 2, baseDelayMs: 10 },
			write: () => (fail ? Promise.reject(new Error('disk full')) : Promise.resolve()),
		});
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(200);
		expect(scheduler.state.status).toBe('failed');
		expect((scheduler.state.lastError as Error).message).toBe('disk full');
		expect(writeBackStateOf(scheduler.state)).toEqual({ status: 'failed', error: 'disk full' });
		expect(scheduler.dirty).toBe(true);
		fail = false;
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(scheduler.state.status).toBe('saved');
		expect(writeBackStateOf(scheduler.state)).toEqual({ status: 'saved', lastSavedAt: 42 });
	});

	it('does not retry without a retry option', async () => {
		const write = vi.fn(() => Promise.reject(new Error('x')));
		const { scheduler } = make({ write });
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(60_000);
		expect(write).toHaveBeenCalledTimes(1);
		expect(scheduler.state.status).toBe('failed');
	});

	it('treats a throwing serialize as a failed try', async () => {
		const { scheduler } = make({
			serialize: () => {
				throw new Error('boom');
			},
		});
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(scheduler.state.status).toBe('failed');
	});

	it('skips a write when serialize resolves null and keeps the status honest', async () => {
		const { scheduler, write } = make({ serialize: () => Promise.resolve(null) });
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(write).not.toHaveBeenCalled();
		expect(scheduler.state.status).toBe('idle');
	});

	it('flush writes now, synchronously for sync hosts, and resolves when nothing is unsaved', async () => {
		const { scheduler, write } = make();
		await scheduler.flush();
		expect(write).not.toHaveBeenCalled();
		scheduler.schedule();
		const done = scheduler.flush();
		expect(write).toHaveBeenCalledTimes(1);
		await done;
		expect(scheduler.state.status).toBe('saved');
		await vi.advanceTimersByTimeAsync(10_000);
		expect(write).toHaveBeenCalledTimes(1);
	});

	it('flush resolves only once the host acknowledged, also a follow-up', async () => {
		const releases: (() => void)[] = [];
		const { scheduler } = make({
			write: () => new Promise<void>((resolve) => releases.push(resolve)),
		});
		scheduler.schedule();
		let resolved = false;
		void scheduler.flush().then(() => (resolved = true));
		await vi.advanceTimersByTimeAsync(0);
		scheduler.schedule();
		releases[0]!();
		await vi.advanceTimersByTimeAsync(0);
		expect(resolved).toBe(false);
		expect(releases).toHaveLength(2);
		releases[1]!();
		await vi.advanceTimersByTimeAsync(0);
		expect(resolved).toBe(true);
	});

	it('flush skips the backoff wait and rejects when the write keeps failing', async () => {
		const write = vi.fn(() => Promise.reject(new Error('down')));
		const { scheduler } = make({ write, retry: { attempts: 3, baseDelayMs: 60_000 } });
		scheduler.schedule();
		const first = scheduler.flush();
		const caught = expect(first).rejects.toThrow('down');
		await vi.advanceTimersByTimeAsync(0);
		await caught;
		expect(write).toHaveBeenCalledTimes(3);
		expect(scheduler.state.status).toBe('failed');
		// Still unsaved: another flush tries again.
		const second = expect(scheduler.flush()).rejects.toThrow('down');
		await vi.advanceTimersByTimeAsync(0);
		await second;
		expect(write).toHaveBeenCalledTimes(6);
	});

	it('cancel drops pending work and ignores a write in flight', async () => {
		const { scheduler, write } = make();
		scheduler.schedule();
		scheduler.cancel();
		await vi.advanceTimersByTimeAsync(1_000);
		expect(write).not.toHaveBeenCalled();
		expect(scheduler.state.status).toBe('idle');
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		expect(write).toHaveBeenCalledTimes(1);
	});

	it('writes a change made after a cancel once the cancelled write has settled', async () => {
		const releases: (() => void)[] = [];
		const { scheduler } = make({
			write: () => new Promise<void>((resolve) => releases.push(resolve)),
		});
		scheduler.schedule();
		await vi.advanceTimersByTimeAsync(100);
		scheduler.cancel();
		scheduler.schedule();
		releases[0]!();
		await vi.advanceTimersByTimeAsync(100);
		expect(releases).toHaveLength(2);
	});

	it('dispose stops everything for good', async () => {
		const { scheduler, write } = make();
		scheduler.schedule();
		scheduler.dispose();
		scheduler.schedule();
		await scheduler.flush();
		await vi.advanceTimersByTimeAsync(10_000);
		expect(write).not.toHaveBeenCalled();
	});

	it('accepts injected timers', () => {
		const fired: (() => void)[] = [];
		const { scheduler, write } = make({
			timers: {
				setTimeout: (cb) => fired.push(cb),
				clearTimeout: () => {},
			},
		});
		scheduler.schedule();
		expect(fired).toHaveLength(1);
		fired[0]!();
		expect(write).toHaveBeenCalledTimes(1);
	});
});

describe('flushWriteBackOnPageClose', () => {
	function targets() {
		const win = new Map<string, () => void>();
		const doc = new Map<string, () => void>();
		const document = {
			visibilityState: 'visible',
			addEventListener: (t: 'visibilitychange', l: () => void) => void doc.set(t, l),
			removeEventListener: (t: 'visibilitychange') => void doc.delete(t),
		};
		return {
			win,
			doc,
			document,
			window: {
				addEventListener: (t: 'pagehide', l: () => void) => void win.set(t, l),
				removeEventListener: (t: 'pagehide') => void win.delete(t),
			},
		};
	}

	it('flushes unsaved work on pagehide and when hidden, and detaches', () => {
		const { scheduler, write } = make();
		const t = targets();
		const detach = flushWriteBackOnPageClose(scheduler, t);
		t.document.visibilityState = 'hidden';
		t.doc.get('visibilitychange')!();
		expect(write).not.toHaveBeenCalled(); // nothing unsaved
		t.document.visibilityState = 'visible';
		scheduler.schedule();
		t.doc.get('visibilitychange')!();
		expect(write).not.toHaveBeenCalled(); // still visible
		t.document.visibilityState = 'hidden';
		t.doc.get('visibilitychange')!();
		expect(write).toHaveBeenCalledTimes(1);
		scheduler.schedule();
		t.win.get('pagehide')!();
		expect(write).toHaveBeenCalledTimes(2);
		detach();
		expect(t.win.size + t.doc.size).toBe(0);
	});

	it('reports a failed flush instead of throwing', async () => {
		const onError = vi.fn();
		const { scheduler } = make({ write: () => Promise.reject(new Error('x')) });
		const t = targets();
		flushWriteBackOnPageClose(scheduler, { ...t, onError });
		scheduler.schedule();
		t.win.get('pagehide')!();
		await vi.advanceTimersByTimeAsync(0);
		expect(onError).toHaveBeenCalledTimes(1);
	});
});
