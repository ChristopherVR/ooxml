// A format-neutral, DOM-free write-back scheduler for an elected writer (a room's owner): a
// trailing debounce, one write in flight, changes made during a write coalesced into one follow-up
// write, a returned promise as the acknowledgement, bounded exponential-backoff retries and a
// status a UI can show. Products supply `serialize` and `write`; see docs/collab-area.md.

/** Where a write-back stands: `idle` until the first change, then `pending`, `saving`, `saved` or `failed`. */
export type WriteBackStatus = 'idle' | 'pending' | 'saving' | 'saved' | 'failed';

/** What {@link WriteBackScheduler.state} and `onStatus` report. */
export interface WriteBackState {
	status: WriteBackStatus;
	/** `now()` when the host last acknowledged a write. */
	lastSavedAt?: number;
	/** The error of the last failed write, once its retries are spent. */
	lastError?: unknown;
}

/** The write-back status as an editor reports it in its collaboration state (JSON friendly). */
export interface CollaborationWriteBackState {
	status: WriteBackStatus;
	lastSavedAt?: number;
	/** The message of the last failure, while `status` is `failed`. */
	error?: string;
}

export interface WriteBackRetry {
	/** Total tries per write, the first included. 1 (the default) never retries. */
	attempts: number;
	/** Wait before the first retry; doubles on each further retry. Default 1000. */
	baseDelayMs?: number;
	/** Upper bound of the wait. Default 30000. */
	maxDelayMs?: number;
}

export interface WriteBackTimers {
	setTimeout(callback: () => void, ms: number): unknown;
	clearTimeout(handle: unknown): void;
}

export interface WriteBackOptions {
	/**
	 * The bytes to store now. Resolve `null` when there is nothing to save yet (the change is not
	 * lost: call `schedule()` again when it is). `isCurrent` turns false once the scheduler was
	 * cancelled or disposed, so a long serialisation can stop early.
	 */
	serialize: (context: {
		isCurrent: () => boolean;
	}) => Uint8Array | null | Promise<Uint8Array | null>;
	/**
	 * Stores the bytes. A returned promise is the acknowledgement: resolving means stored, rejecting
	 * counts as a failed try. A plain return (void) is treated as acknowledged at once.
	 */
	write: (bytes: Uint8Array) => void | Promise<void>;
	/** Quiet time after the last change. Default 5000; invalid values fall back to it. */
	debounceMs?: number;
	retry?: WriteBackRetry;
	onStatus?: (state: WriteBackState) => void;
	/** Clock for `lastSavedAt`. Default `Date.now`. */
	now?: () => number;
	/** Timers, injectable for tests. Default the global ones, looked up at call time. */
	timers?: WriteBackTimers;
}

export interface WriteBackScheduler {
	/** Notes a change: (re)starts the quiet time, or queues one follow-up write during a write. */
	schedule(): void;
	/**
	 * Writes now and resolves once the host acknowledged it (and anything changed meanwhile).
	 * Resolves at once when nothing is unsaved; rejects with the last error when the write failed
	 * after its retries. When `serialize` and `write` return synchronously, the write has been
	 * handed over before `flush` returns.
	 */
	flush(): Promise<void>;
	/** Drops pending work and ignores the result of a write in flight; the scheduler stays usable. */
	cancel(): void;
	/** Cancels and stops for good: later calls do nothing. */
	dispose(): void;
	readonly state: WriteBackState;
	/** True while a change waits for its quiet time or a write (or its retry) is under way. */
	readonly dirty: boolean;
}

export const DEFAULT_WRITE_BACK_DEBOUNCE_MS = 5_000;
const DEFAULT_RETRY_BASE_MS = 1_000;
const DEFAULT_RETRY_MAX_MS = 30_000;

const globalTimers: WriteBackTimers = {
	setTimeout: (callback, ms) => globalThis.setTimeout(callback, ms),
	clearTimeout: (handle) => globalThis.clearTimeout(handle as ReturnType<typeof setTimeout>),
};

const isThenable = <T>(value: unknown): value is PromiseLike<T> =>
	typeof value === 'object' &&
	value !== null &&
	typeof (value as PromiseLike<T>).then === 'function';

/** The status of `state` for a collaboration state; the error becomes its message. */
export function writeBackStateOf(state: WriteBackState): CollaborationWriteBackState {
	const out: CollaborationWriteBackState = { status: state.status };
	if (state.lastSavedAt !== undefined) out.lastSavedAt = state.lastSavedAt;
	if (state.status === 'failed' && state.lastError !== undefined) {
		out.error =
			state.lastError instanceof Error ? state.lastError.message : String(state.lastError);
	}
	return out;
}

export function createWriteBackScheduler(options: WriteBackOptions): WriteBackScheduler {
	const timers = options.timers ?? globalTimers;
	const now = options.now ?? Date.now;
	const configured = options.debounceMs;
	const debounceMs =
		configured !== undefined && Number.isFinite(configured) && configured >= 0
			? configured
			: DEFAULT_WRITE_BACK_DEBOUNCE_MS;
	const attempts = Math.max(1, Math.floor(options.retry?.attempts ?? 1) || 1);
	const baseDelay = Math.max(0, options.retry?.baseDelayMs ?? DEFAULT_RETRY_BASE_MS);
	const maxDelay = Math.max(baseDelay, options.retry?.maxDelayMs ?? DEFAULT_RETRY_MAX_MS);

	let state: WriteBackState = { status: 'idle' };
	let generation = 0;
	let disposed = false;
	let dirty = false;
	let timer: unknown = null;
	let running: Promise<void> | null = null;
	let wake: (() => void) | null = null;
	let ended = false;
	/** Flushes under way: a flush does not wait out a retry backoff. */
	let rushing = 0;

	const patch = (status: WriteBackStatus, extra: Partial<WriteBackState> = {}): void => {
		const next: WriteBackState = { status };
		const saved = extra.lastSavedAt ?? state.lastSavedAt;
		if (saved !== undefined) next.lastSavedAt = saved;
		if (extra.lastError !== undefined) next.lastError = extra.lastError;
		state = next;
		options.onStatus?.(state);
	};
	const clearTimer = (): void => {
		if (timer !== null) timers.clearTimeout(timer);
		timer = null;
	};
	const arm = (): void => {
		clearTimer();
		timer = timers.setTimeout(() => {
			timer = null;
			void start();
		}, debounceMs);
		if (state.status !== 'pending') patch('pending');
	};
	const sleep = (ms: number): Promise<void> =>
		new Promise<void>((resolve) => {
			if (rushing > 0) return resolve();
			const handle = timers.setTimeout(done, ms);
			function done(): void {
				timers.clearTimeout(handle);
				wake = null;
				resolve();
			}
			wake = done;
		});
	const settledStatus = (): WriteBackStatus => (state.lastSavedAt === undefined ? 'idle' : 'saved');

	async function run(): Promise<void> {
		const gen = generation;
		const isCurrent = (): boolean => gen === generation;
		patch('saving');
		try {
			for (let attempt = 1; ; attempt++) {
				dirty = false;
				try {
					let bytes = options.serialize({ isCurrent });
					if (isThenable<Uint8Array | null>(bytes)) bytes = await bytes;
					if (!isCurrent()) return;
					if (bytes === null || bytes === undefined) {
						if (dirty) arm();
						else patch(settledStatus());
						return;
					}
					const acknowledged = options.write(bytes);
					if (isThenable<void>(acknowledged)) await acknowledged;
					if (!isCurrent()) return;
					const lastSavedAt = now();
					if (dirty) {
						patch('pending', { lastSavedAt });
						arm();
					} else {
						patch('saved', { lastSavedAt });
					}
					return;
				} catch (error) {
					if (!isCurrent()) return;
					if (attempt >= attempts) {
						dirty = true;
						patch('failed', { lastError: error });
						return;
					}
					await sleep(Math.min(maxDelay, baseDelay * 2 ** (attempt - 1)));
					if (!isCurrent()) return;
				}
			}
		} finally {
			// Synchronous with the last step, so a void write leaves nothing in flight.
			ended = true;
			running = null;
			// A change made after a cancel, while the cancelled write was still settling.
			if (dirty && timer === null && !disposed && state.status !== 'failed') arm();
		}
	}

	function start(): Promise<void> {
		if (running) return running;
		ended = false;
		const promise = run();
		// A run that finished synchronously (a void write) leaves nothing in flight.
		if (!ended) running = promise;
		return promise;
	}

	function cancel(): void {
		generation++;
		clearTimer();
		dirty = false;
		const waiting = wake;
		wake = null;
		waiting?.();
		if (state.status === 'pending' || state.status === 'saving') patch(settledStatus());
	}

	return {
		schedule(): void {
			if (disposed) return;
			dirty = true;
			// During a write the follow-up is armed when it finishes.
			if (running) return;
			arm();
		},
		async flush(): Promise<void> {
			if (disposed) return;
			clearTimer();
			rushing++;
			try {
				// A retry waiting out its backoff goes now.
				wake?.();
				for (;;) {
					if (running) await running;
					else if (dirty) await start();
					else return;
					if (disposed) return;
					if (state.status === 'failed') throw state.lastError;
					clearTimer();
				}
			} finally {
				rushing--;
			}
		},
		cancel,
		dispose(): void {
			if (disposed) return;
			cancel();
			disposed = true;
		},
		get state(): WriteBackState {
			return state;
		},
		get dirty(): boolean {
			return dirty;
		},
	};
}
