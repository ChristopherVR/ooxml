export interface RecoveryTimers {
	set(handler: () => void, delayMs: number): unknown;
	clear(id: unknown): void;
}

/** Timer ownership and serialized recovery writes; products own cadence and snapshot policy. */
export function createRecoveryScheduler(options: {
	tick(): void;
	timers?: RecoveryTimers;
	/** True only when injected timers.set already repeats (for example, setInterval). */
	repeating?: boolean;
}) {
	const timers = options.timers ?? {
		set: (handler: () => void, ms: number) => setTimeout(handler, ms),
		clear: (id: unknown) => clearTimeout(id as ReturnType<typeof setTimeout>),
	};
	let timer: unknown;
	let armed = false;
	let epoch = 0;
	let disposed = false;
	let inFlight: Promise<unknown> | undefined;
	const clear = (): void => {
		++epoch;
		if (armed) timers.clear(timer);
		armed = false;
	};
	return {
		clear,
		arm(delayMs: number): void {
			clear();
			if (disposed) return;
			armed = true;
			const generation = epoch;
			timer = timers.set(() => {
				if (disposed || epoch !== generation) return;
				if (!options.repeating) armed = false;
				options.tick();
			}, delayMs);
		},
		/** Re-read live product state only once earlier writes have completed. */
		async run<T>(write: () => Promise<T>): Promise<T | undefined> {
			while (inFlight) {
				try {
					await inFlight;
				} catch {
					/* A failed write does not block its successor. */
				}
			}
			if (disposed) return;
			const pending = write();
			inFlight = pending;
			try {
				return await pending;
			} finally {
				inFlight = undefined;
			}
		},
		dispose(): void {
			disposed = true;
			clear();
		},
	};
}
