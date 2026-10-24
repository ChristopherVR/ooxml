/**
 * One write queue per handler for its theme parts (`render/theme-history`).
 *
 * Each theme part write reads, edits and writes the part back, so the writes on one handler run one
 * after another ({@link queueThemeWrite}), and every save waits for the ones queued before it
 * ({@link settleThemeWrites}, which `saveDeckWithPassword` calls).
 *
 * @module render/theme-write-queue
 */

/** Per handler, the end of its queue: settles once the last queued write has landed or failed. */
const pendingWrites = new WeakMap<object, Promise<void>>();

function startWrite<T>(write: () => Promise<T>): Promise<T> {
	try {
		return write();
	} catch (error) {
		return Promise.reject(error);
	}
}

/**
 * Run `write` on `handler`'s theme parts after every write queued on it before. An undo landing
 * while a pick is still writing (or two undos in a row) waits for it instead of interleaving.
 * With nothing queued the write starts at once, exactly as an unqueued one would.
 */
export function queueThemeWrite<T>(handler: object, write: () => Promise<T>): Promise<T> {
	const previous = pendingWrites.get(handler);
	const run = previous ? previous.then(write) : startWrite(write);
	const settled: Promise<void> = run.then(
		() => undefined,
		() => undefined,
	);
	const end = settled.then(() => {
		if (pendingWrites.get(handler) === end) {
			pendingWrites.delete(handler);
		}
	});
	pendingWrites.set(handler, end);
	return run;
}

/**
 * Resolves once every theme part write queued on `handler` so far has landed (or failed). A save
 * waits for this, or it writes a theme an undo or a pick is still replacing.
 */
export function settleThemeWrites(handler: object | null | undefined): Promise<void> {
	return (handler && pendingWrites.get(handler)) || Promise.resolve();
}
