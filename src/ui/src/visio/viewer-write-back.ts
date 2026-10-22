import type { ViewerController } from './controller';
import type { VisioWriteBackOptions } from './collaboration-types';

/** Default quiet time before a write-back (the same as PowerPoint's collaboration config). */
export const WRITE_BACK_DEBOUNCE_MS = 5_000;

/**
 * Elected-writer write-back of one shared drawing: only the room's owner persists, so concurrent
 * collaborators do not race. A trailing debounce over local and remote changes saves the drawing
 * and hands the bytes to `onWriteBack`; leaving flushes a pending save once. The same scheduling
 * policy as PowerPoint's write-back, kept here because its serializer and session types are
 * PowerPoint's. The callback's result is not a storage acknowledgement and is never retried.
 */
export class ViewerWriteBack {
	readonly #stops: (() => void)[] = [];
	#timer: ReturnType<typeof setTimeout> | null = null;
	#disposed = false;
	constructor(
		private readonly controller: ViewerController,
		private readonly options: VisioWriteBackOptions,
		/** True while this window is the owner and may write. */
		private readonly eligible: () => boolean,
	) {
		if (!options.onWriteBack) return;
		this.#stops.push(
			controller.onEvent((name) => {
				if (name === 'document-change' || name === 'document-load') this.#schedule();
			}),
		);
	}
	/** Whether a save is waiting for its quiet time. */
	get pending(): boolean {
		return this.#timer !== null;
	}
	/** Fires a pending write-back now. */
	flush(): void {
		if (this.#timer === null) return;
		this.#cancel();
		this.#fire();
	}
	/** Flushes a pending write-back and stops listening. */
	dispose(): void {
		if (this.#disposed) return;
		this.flush();
		this.#disposed = true;
		for (const stop of this.#stops.splice(0)) stop();
	}
	#cancel(): void {
		if (this.#timer !== null) clearTimeout(this.#timer);
		this.#timer = null;
	}
	#schedule(): void {
		if (this.#disposed || !this.eligible()) return;
		this.#cancel();
		const ms = this.options.writeBackDebounceMs;
		this.#timer = setTimeout(
			() => {
				this.#timer = null;
				this.#fire(true);
			},
			ms !== undefined && Number.isFinite(ms) && ms >= 0 ? ms : WRITE_BACK_DEBOUNCE_MS,
		);
	}
	#fire(retryWhenBusy = false): void {
		const { onWriteBack } = this.options;
		if (!onWriteBack || !this.eligible()) return;
		const { edit, loading } = this.controller.state;
		if (edit.busy || loading) {
			// A save in flight: look again after another quiet period instead of exporting half-way.
			if (retryWhenBusy) this.#schedule();
			return;
		}
		let bytes: Uint8Array;
		try {
			bytes = this.controller.exportVsdx().bytes;
		} catch {
			return; // nothing editable to save (no source package)
		}
		try {
			const result: unknown = onWriteBack(bytes);
			if (result instanceof Promise) result.catch(() => {});
		} catch {
			/* the host's failure is not the room's */
		}
	}
}
