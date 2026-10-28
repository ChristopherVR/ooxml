/**
 * The recovery autosave engine every PowerPoint binding adapts.
 *
 * Until this controller each binding owned an engine of its own: React and
 * Angular polled on a `setInterval`, Vue, Svelte and vanilla debounced on the
 * edit, and each reported a status shape of its own. There is now ONE engine,
 * following PowerPoint's AutoRecover: while there are unsaved edits a recovery
 * snapshot is written every interval (File > Options > Save > "Save
 * AutoRecover information every N minutes", two minutes by default here, see
 * `resolveAutosaveIntervalMs`), and nothing is written while the document is
 * unchanged. Concretely:
 *
 *  - the first unsaved edit starts the interval; further edits do not push the
 *    deadline out (`nextAutosaveDelayMs`), so a snapshot lands no later than one
 *    interval after the first unsaved edit, and no more often than once per
 *    interval;
 *  - a snapshot that lands clears the outstanding edits; an edit made while it
 *    was being written starts the next interval;
 *  - an explicit {@link AutosaveController.saveNow} is a request, not a poll:
 *    it writes whenever there is a file key, dirty or not;
 *  - a Design gallery hover preview ({@link AutosaveControllerDeps.isHeld}) is
 *    never written: the timer tries again one interval later.
 *
 * Whether autosave runs at all is the shared `resolveAutosaveActivation`
 * verdict, read through {@link AutosaveControllerDeps.getActivation}; the
 * engine never owns a raw "enabled" flag. Each binding tells the engine about
 * edits in its own reactive idiom, either directly ({@link
 * AutosaveController.markEdited}) or by handing it the editor's dirty flag and
 * the values a snapshot is built from ({@link AutosaveController.observe}).
 *
 * @module render/controller/autosave-controller
 */

import type { DocumentStorage } from 'ooxml-core/embedding';

import type { AutosaveDisabledReason } from '../autosave-policy';
import { nextAutosaveDelayMs } from '../autosave-policy';
import { saveAutosaveSnapshot } from '../autosave-store';
import { autosaveSnapshotMark, autosaveSourcesChanged } from '../autosave-tick';
import type { AutosaveSnapshotMark } from '../autosave-tick';
import type { InlineTextDraft } from '../autosave-inline-draft';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';
import { createRecoveryScheduler } from '../../../recovery-scheduler';

/** The one autosave status every binding renders. */
export type PptxAutosaveStatus =
	| { readonly state: 'idle' }
	| { readonly state: 'disabled'; readonly reason: string }
	| { readonly state: 'saving' }
	| { readonly state: 'saved'; readonly timestamp: number }
	| { readonly state: 'error'; readonly message: string };

/** The status without its payload, for bindings that render a bare string. */
export type PptxAutosaveStatusKind = PptxAutosaveStatus['state'];

/** What the chrome reads from the engine. */
export interface AutosaveControllerState {
	readonly status: PptxAutosaveStatus;
	/** Edits not yet captured by a recovery snapshot. */
	readonly dirty: boolean;
	/** When the last snapshot landed (epoch ms), or null. */
	readonly lastSavedAt: number | null;
}

/** Timer seam; defaults to the global `setTimeout` / `clearTimeout`. */
export interface AutosaveTimers {
	setTimeout(handler: () => void, ms: number): unknown;
	clearTimeout(id: unknown): void;
}

/** The part of the shared activation verdict the engine needs. */
export interface AutosaveEngineActivation {
	readonly active: boolean;
	readonly reason?: AutosaveDisabledReason | undefined;
}

export interface AutosaveControllerDeps {
	/** The shared activation verdict, read live (see `resolveAutosaveActivation`). */
	getActivation(): AutosaveEngineActivation;
	/** The snapshot key (the host's file path or name). Inert without one. */
	getFilePath(): string | undefined;
	/** The resolved cadence (`resolveAutosaveIntervalMs`), read each time the timer is armed. */
	getIntervalMs(): number;
	/**
	 * Serialise the deck for recovery: the plaintext `recovery-snapshot` bytes,
	 * never the user's encrypted save. `null` skips the write.
	 */
	serialize(draft: InlineTextDraft | null): Promise<Uint8Array | null>;
	/**
	 * The element whose inline editor is open, if any. A draft is only handed
	 * to {@link serialize} while its element is still being edited, so a
	 * cancelled edit never reaches a snapshot.
	 */
	getInlineEditingId?(): string | null | undefined;
	/** Host `storage`: snapshots go to its recovery methods when it has them. */
	getStorage?(): DocumentStorage | undefined;
	/** True while a hover preview stands in for the document; a due write waits. */
	isHeld?(): boolean;
	/** Called with the bytes of each snapshot written (the host's autosave event). */
	onSaved?(bytes: Uint8Array): void;
	/** Message for a failure that is not an `Error`. */
	errorMessage?(error: unknown): string;
	timers?: AutosaveTimers;
	now?(): number;
	/** Where snapshots go; defaults to the shared `saveAutosaveSnapshot`. */
	writeSnapshot?(
		filePath: string,
		bytes: Uint8Array,
		storage: DocumentStorage | undefined,
	): Promise<unknown>;
}

export interface AutosaveController extends PptxController<AutosaveControllerState> {
	/** An edit landed: mark the document dirty and start the interval. */
	markEdited(): void;
	/**
	 * Inline text typed but not committed yet (see `reportInlineTextDraft`):
	 * a changed draft is an edit, and the next snapshot includes it. `null`
	 * clears it (the editor closed).
	 */
	noteInlineDraft(draft: InlineTextDraft | null): void;
	/**
	 * Nothing is outstanding any more (a deck was loaded, or the user saved).
	 * `saved: true` also shows "saved" (a real save captured everything).
	 */
	markClean(options?: { saved?: boolean }): void;
	/**
	 * Feed the editor's own dirty flag and the values a snapshot is built from
	 * (reassigned, never mutated). Clean clears; dirty with any changed value is
	 * an edit; dirty with nothing changed is not.
	 */
	observe(dirty: boolean, sources: readonly unknown[]): void;
	/** Write a snapshot now, dirty or not (still needs a file key). */
	saveNow(): Promise<void>;
	/** Re-read the activation, the file key and the cadence. */
	refresh(): void;
	destroy(): void;
}

const defaultTimers: AutosaveTimers = {
	setTimeout: (handler, ms) => setTimeout(handler, ms),
	clearTimeout: (id) => clearTimeout(id as ReturnType<typeof setTimeout>),
};

function disabledReason(
	activation: AutosaveEngineActivation,
	filePath: string | undefined,
): string {
	return activation.reason ?? (filePath ? 'autosave_toggle_off' : 'no_file_path');
}

/** Build the recovery autosave engine (see the module note). */
export function createAutosaveController(deps: AutosaveControllerDeps): AutosaveController {
	const timers = deps.timers ?? defaultTimers;
	const now = deps.now ?? Date.now;
	const write = deps.writeSnapshot ?? saveAutosaveSnapshot;

	const isActive = (): boolean => deps.getActivation().active && Boolean(deps.getFilePath());
	const idleOrDisabled = (): PptxAutosaveStatus =>
		isActive()
			? { state: 'idle' }
			: { state: 'disabled', reason: disabledReason(deps.getActivation(), deps.getFilePath()) };

	const store = createControllerStore<AutosaveControllerState>({
		status: idleOrDisabled(),
		dirty: false,
		lastSavedAt: null,
	});
	const patch = (next: Partial<AutosaveControllerState>): void => {
		const current = store.getSnapshot();
		const changed = (Object.keys(next) as (keyof AutosaveControllerState)[]).some(
			(key) => next[key] !== current[key],
		);
		if (changed) {
			store.set({ ...current, ...next });
		}
	};

	let active = isActive();
	let timer: unknown = null;
	let firstDirtyAt: number | null = null;
	/** Set when an edit lands while a snapshot is being written. */
	let editedDuringSave: number | null = null;
	let inFlight: Promise<void> | null = null;
	let observed: AutosaveSnapshotMark | undefined;
	let observedDirty = false;
	let inlineDraft: InlineTextDraft | null = null;
	/** The draft still belongs to an open editor (or the binding does not say). */
	const liveDraft = (): InlineTextDraft | null => {
		if (!inlineDraft) {
			return null;
		}
		const editing = deps.getInlineEditingId?.();
		if (editing !== undefined && editing !== inlineDraft.elementId) {
			inlineDraft = null;
		}
		return inlineDraft;
	};
	let destroyed = false;
	const scheduler = createRecoveryScheduler({
		timers: {
			set: (handler, ms) => timers.setTimeout(handler, ms),
			clear: (id) => timers.clearTimeout(id),
		},
		tick: () => {
			timer = null;
			if (destroyed || !isActive() || !store.getSnapshot().dirty) return;
			if (deps.isHeld?.()) {
				firstDirtyAt = now();
				schedule();
				return;
			}
			void saveNow();
		},
	});

	const clearTimer = (): void => {
		scheduler.clear();
		timer = null;
	};

	const schedule = (): void => {
		clearTimer();
		if (destroyed || !active) {
			return;
		}
		const delay = nextAutosaveDelayMs({
			intervalMs: deps.getIntervalMs(),
			firstDirtyAt,
			now: now(),
		});
		timer = true;
		scheduler.arm(delay);
	};

	const run = async (): Promise<void> => {
		const filePath = deps.getFilePath();
		if (!filePath || destroyed) {
			return;
		}
		editedDuringSave = null;
		clearTimer();
		patch({ status: { state: 'saving' } });
		try {
			const bytes = await deps.serialize(liveDraft());
			if (destroyed) return;
			if (!bytes) {
				if (!destroyed) {
					patch({ status: idleOrDisabled() });
				}
				return;
			}
			await write(filePath, bytes, deps.getStorage?.());
			if (destroyed) {
				return;
			}
			const savedAt = now();
			const outstanding = editedDuringSave;
			firstDirtyAt = outstanding;
			patch({
				status: { state: 'saved', timestamp: savedAt },
				dirty: outstanding !== null,
				lastSavedAt: savedAt,
			});
			if (outstanding !== null) {
				schedule();
			}
			deps.onSaved?.(bytes);
		} catch (error) {
			if (!destroyed) {
				patch({
					status: {
						state: 'error',
						message:
							error instanceof Error
								? error.message
								: (deps.errorMessage?.(error) ?? 'Autosave failed'),
					},
				});
			}
		} finally {
			editedDuringSave = null;
		}
	};

	const saveNow = async (): Promise<void> => {
		// Requests queue behind a write in flight rather than overlapping it.
		while (inFlight) {
			await inFlight;
		}
		const pending = scheduler.run(run).then(() => {});
		inFlight = pending;
		try {
			await pending;
		} finally {
			inFlight = null;
		}
	};

	const markEdited = (): void => {
		if (destroyed) {
			return;
		}
		const at = now();
		if (inFlight) {
			editedDuringSave ??= at;
		}
		firstDirtyAt ??= at;
		patch({ dirty: true });
		if (active && !inFlight) {
			schedule();
		}
	};

	const markClean = (options?: { saved?: boolean }): void => {
		clearTimer();
		firstDirtyAt = null;
		editedDuringSave = null;
		patch(
			options?.saved && active
				? { dirty: false, status: { state: 'saved', timestamp: now() } }
				: { dirty: false },
		);
	};

	const refresh = (): void => {
		if (destroyed) {
			return;
		}
		const next = isActive();
		const was = active;
		active = next;
		if (!next) {
			clearTimer();
			patch({ status: idleOrDisabled() });
			return;
		}
		if (store.getSnapshot().status.state === 'disabled') {
			patch({ status: { state: 'idle' } });
		}
		if (!store.getSnapshot().dirty || inFlight) {
			return;
		}
		if (!was) {
			// Autosave was off while these edits happened: the interval starts now.
			firstDirtyAt = now();
			schedule();
		} else if (timer !== null) {
			// The cadence may have changed: re-arm the pending timer with it.
			schedule();
		}
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		markEdited,
		noteInlineDraft(draft) {
			const previous = inlineDraft;
			inlineDraft = draft;
			if (
				draft &&
				(previous?.elementId !== draft.elementId ||
					previous.slideId !== draft.slideId ||
					previous.text !== draft.text)
			) {
				markEdited();
			}
		},
		markClean,
		observe(dirty, sources) {
			const changed = autosaveSourcesChanged(observed, sources);
			const turnedDirty = dirty && !observedDirty;
			observed = autosaveSnapshotMark('', sources);
			observedDirty = dirty;
			if (!dirty) {
				// The editor is clean, but typing still open in an inline editor is not.
				if (store.getSnapshot().dirty && !liveDraft()) {
					markClean();
				}
				return;
			}
			if (turnedDirty || changed) {
				markEdited();
			}
		},
		saveNow,
		refresh,
		destroy() {
			destroyed = true;
			clearTimer();
			scheduler.dispose();
		},
	};
}
