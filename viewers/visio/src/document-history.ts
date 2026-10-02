/** Source snapshots only. No document model is serialized by the viewer. */
export interface EditDiagnostic {
	readonly code: string;
	readonly message: string;
}
export interface VsdxExportResult {
	bytes: Uint8Array;
	dirty: boolean;
	diagnostics: readonly EditDiagnostic[];
}
export interface ViewerEditState {
	readonly sourceAvailable: boolean;
	readonly busy: boolean;
	readonly dirty: boolean;
	readonly canUndo: boolean;
	readonly canRedo: boolean;
	readonly historyTruncated: boolean;
	readonly error: Error | null;
	readonly diagnostics: readonly EditDiagnostic[];
}
export const EMPTY_EDIT_STATE: ViewerEditState = Object.freeze({
	sourceAvailable: false,
	busy: false,
	dirty: false,
	canUndo: false,
	canRedo: false,
	historyTruncated: false,
	error: null,
	diagnostics: Object.freeze([]),
});
export const EDIT_HISTORY_LIMITS = Object.freeze({ maxBytes: 96 * 1024 * 1024, maxEntries: 20 });
export interface SourceSnapshot {
	readonly bytes: Uint8Array;
	readonly diagnostics: readonly EditDiagnostic[];
}
/** Private controller-owned bytes; transient worker copies and models are separate from this budget. */
export class DocumentHistory {
	readonly original: SourceSnapshot;
	#entries: SourceSnapshot[];
	#index = 0;
	#truncated = false;
	constructor(bytes: Uint8Array) {
		this.original = { bytes: Uint8Array.from(bytes), diagnostics: Object.freeze([]) };
		this.#entries = [this.original];
	}
	get current(): SourceSnapshot {
		return this.#entries[this.#index]!;
	}
	get undoTarget(): SourceSnapshot | undefined {
		return this.#entries[this.#index - 1];
	}
	get redoTarget(): SourceSnapshot | undefined {
		return this.#entries[this.#index + 1];
	}
	get state(): ViewerEditState {
		return Object.freeze({
			...EMPTY_EDIT_STATE,
			sourceAvailable: true,
			dirty: this.current !== this.original,
			canUndo: !!this.undoTarget,
			canRedo: !!this.redoTarget,
			historyTruncated: this.#truncated,
			diagnostics: this.current.diagnostics,
		});
	}
	append(bytes: Uint8Array, diagnostics: readonly EditDiagnostic[]): void {
		const snapshot = {
			bytes: Uint8Array.from(bytes),
			diagnostics: Object.freeze(
				diagnostics.map((note) => Object.freeze({ code: note.code, message: note.message })),
			),
		};
		this.#entries = this.#entries.slice(0, this.#index + 1);
		this.#entries.push(snapshot);
		this.#index++;
		const retained = () => new Set([this.original, ...this.#entries]);
		while (
			this.#entries.length > EDIT_HISTORY_LIMITS.maxEntries ||
			[...retained()].reduce((sum, entry) => sum + entry.bytes.byteLength, 0) >
				EDIT_HISTORY_LIMITS.maxBytes
		) {
			// Original remains privately pinned even when no longer reachable through undo.
			if (this.#index === 0) throw new Error('Source snapshots exceed the retained history limit.');
			this.#entries.shift();
			this.#index--;
			this.#truncated = true;
		}
	}
	move(target: SourceSnapshot): void {
		const index = this.#entries.indexOf(target);
		if (index < 0 || Math.abs(index - this.#index) !== 1)
			throw new Error('The history target is stale.');
		this.#index = index;
	}
	export(): VsdxExportResult {
		return {
			bytes: Uint8Array.from(this.current.bytes),
			dirty: this.current !== this.original,
			diagnostics: this.current.diagnostics,
		};
	}
}
