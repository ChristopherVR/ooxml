import {
	createWriteBackScheduler,
	DEFAULT_WRITE_BACK_DEBOUNCE_MS,
	flushWriteBackOnPageClose,
	writeBackStateOf,
	type CollaborationWriteBackState,
	type WriteBackScheduler,
} from 'ooxml-core/collab';
import type { ViewerController } from './controller';
import type { VisioWriteBackOptions } from './collaboration-types';

/** Default quiet time before a write-back (the same as PowerPoint's collaboration config). */
export const WRITE_BACK_DEBOUNCE_MS = DEFAULT_WRITE_BACK_DEBOUNCE_MS;

/**
 * Elected-writer write-back of one shared drawing: only the room's owner persists, so concurrent
 * collaborators do not race. The scheduling policy (trailing debounce, one write in flight,
 * acknowledgement by a returned promise, bounded retries) is `createWriteBackScheduler` of
 * `ooxml-core/collab`, shared with PowerPoint. Leaving, `pagehide` and the page turning hidden
 * flush a pending save; the browser does not wait for an asynchronous write then (see the
 * collaboration docs for `sendBeacon` and `keepalive`).
 */
export class ViewerWriteBack {
	readonly #stops: (() => void)[] = [];
	readonly #scheduler: WriteBackScheduler | null;
	#disposed = false;
	constructor(
		private readonly controller: ViewerController,
		options: VisioWriteBackOptions,
		/** True while this window is the owner and may write. */
		private readonly eligible: () => boolean,
		/** Called whenever {@link state} changed. */
		onStatus: () => void = () => {},
	) {
		const { onWriteBack } = options;
		if (!onWriteBack) {
			this.#scheduler = null;
			return;
		}
		this.#scheduler = createWriteBackScheduler({
			debounceMs: options.writeBackDebounceMs ?? DEFAULT_WRITE_BACK_DEBOUNCE_MS,
			...(options.writeBackRetry ? { retry: options.writeBackRetry } : {}),
			onStatus: () => onStatus(),
			serialize: () => this.#serialize(),
			write: (bytes) => onWriteBack(bytes),
		});
		const self = this;
		this.#stops.push(
			controller.onEvent((name) => {
				if (name === 'document-change' || name === 'document-load') this.#schedule();
			}),
			flushWriteBackOnPageClose({
				get dirty() {
					return self.pending && self.eligible();
				},
				flush: () => this.flush(),
			}),
		);
	}
	/** Whether a save is waiting for its quiet time or being written. */
	get pending(): boolean {
		return this.#scheduler?.dirty ?? false;
	}
	/** The write-back status for the collaboration state; undefined without `onWriteBack`. */
	get state(): CollaborationWriteBackState | undefined {
		return this.#scheduler ? writeBackStateOf(this.#scheduler.state) : undefined;
	}
	/** Writes a pending change now; resolves when the host acknowledged it. */
	flush(): Promise<void> {
		return this.#scheduler?.flush() ?? Promise.resolve();
	}
	/** Flushes a pending write-back and stops listening. */
	dispose(): void {
		if (this.#disposed) return;
		const scheduler = this.#scheduler;
		const done = this.flush().catch(() => {});
		this.#disposed = true;
		for (const stop of this.#stops.splice(0)) stop();
		// An asynchronous write and its retries get the chance to finish first.
		void done.then(() => scheduler?.dispose());
	}
	#schedule(): void {
		if (this.#disposed || !this.eligible()) return;
		this.#scheduler?.schedule();
	}
	#serialize(): Uint8Array | null {
		if (!this.eligible()) return null;
		const { edit, loading } = this.controller.state;
		if (edit.busy || loading) {
			// A save in flight: look again after another quiet period instead of exporting half-way.
			this.#schedule();
			return null;
		}
		try {
			return this.controller.exportVsdx().bytes;
		} catch {
			return null; // nothing editable to save (no source package)
		}
	}
}
