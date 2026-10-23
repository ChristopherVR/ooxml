import type {
	AutosaveController as SharedAutosaveController,
	AutosaveControllerState,
	DeckSaveIntent,
	DocumentStorage,
	InlineTextDraft,
	PptxAutosaveStatusKind,
} from 'ooxml-ui/pptx';
import {
	applyInlineTextDraft,
	createAutosaveController,
	createPreviewEditGuard,
	recoverySnapshotIntent,
	saveAutosaveSnapshot,
	saveDeckWithPassword,
} from 'ooxml-ui/pptx';
import type {
	PptxHandoutMaster,
	PptxHandler,
	PptxNotesMaster,
	PptxSection,
	PptxSlide,
	PptxSlideMaster,
} from 'pptx-viewer-core';
import { untrack } from 'svelte';

import { controllerState } from './controller-state';
import type { ControllerState } from './controller-state';

/**
 * autosave.svelte.ts: the Svelte adapter of the shared recovery autosave engine
 * (`createAutosaveController` in `ooxml-ui/pptx`).
 *
 * The engine follows PowerPoint's AutoRecover: one snapshot an interval after
 * the first unsaved edit, written to the shared recovery store (IndexedDB, or
 * the host's recovery methods) and nothing while the deck is unchanged. This
 * class only tells it about edits (telling a load and a hover preview apart
 * from a real edit) and serialises the deck; `status` / `isDirty` /
 * `lastSavedAt` are the engine's state, read reactively.
 */

/** The `state` of the one shared `PptxAutosaveStatus`. */
export type AutosaveStatus = PptxAutosaveStatusKind;

export interface AutosaveDeps {
	/**
	 * Master on/off: the resolved activation (`resolveAutosaveActivation` in
	 * shared), which folds the host policy, the user's toggle, editability and
	 * the collaboration read-only veto into one answer.
	 */
	getEnabled: () => boolean;
	/** Debounce window in milliseconds, and the ceiling on deferring a save. */
	getIntervalMs: () => number;
	/** IndexedDB record key (host `filePath`); autosave is disabled without one. */
	getFilePath: () => string | undefined;
	/** Host `storage` prop: recovery snapshots go to its recovery methods when it has them. */
	getStorage?: () => DocumentStorage | undefined;
	/** The current editable slides (watched for edits). */
	getSlides: () => PptxSlide[];
	/**
	 * True while a Design gallery hover preview shows over the slides, and the identities of the
	 * stores a preview swaps: the effect then tells an edit from a preview and its restore.
	 */
	getPreviewHeld?: () => boolean;
	getDocumentKey?: () => readonly unknown[];
	getSlideMasters?: () => PptxSlideMaster[];
	getNotesMaster?: () => PptxNotesMaster | undefined;
	getHandoutMaster?: () => PptxHandoutMaster | undefined;
	getSections?: () => PptxSection[];
	/** The live core handler used to serialize slides to `.pptx` bytes. */
	getHandler: () => PptxHandler | null;
	/**
	 * The Protect-Presentation state, forwarded to the shared save decision.
	 * Optional, and deliberately inert: a recovery snapshot is written in the
	 * clear whatever it says (see `deck-save-encryption` in
	 * `pptx-viewer-shared`). It is threaded through so the snapshot path uses
	 * the same one decision function as every real save.
	 */
	getSaveIntent?: () => DeckSaveIntent;
	/**
	 * Monotonic load counter: a change means a fresh presentation was seeded (not
	 * a user edit), so the watcher clears dirty instead of arming a save.
	 */
	getLoadCount: () => number;
	/**
	 * Monotonic SEED counter, bumped in the same synchronous block that installs
	 * the loaded slides (`EditorState.seedNonce`).
	 *
	 * `getLoadCount` alone is not enough, because the loader bumps it in an
	 * earlier flush than the effect that copies its slides into the editor: this
	 * watcher ran once on the count change (old slides, so nothing to do) and
	 * then a second time on the slide reassignment, with the count already
	 * settled - indistinguishable from an edit. Merely OPENING a deck therefore
	 * marked it dirty and wrote a crash-recovery snapshot, so the next visit
	 * offered to recover unsaved changes that never existed.
	 */
	getSeedNonce?: () => number;
	/** The element whose inline editor is open; a closed editor's draft is dropped. */
	getInlineEditingId?: () => string | null;
	/** Host callback with the freshly-serialized bytes on each successful save. */
	onSaved?: (bytes: Uint8Array) => void;
}

export class AutosaveController {
	readonly #deps: AutosaveDeps;
	readonly #engine: SharedAutosaveController;
	readonly #state: ControllerState<AutosaveControllerState>;
	#lastLoadCount = 0;
	#lastSeedNonce = 0;
	#started = false;
	readonly #previewGuard = createPreviewEditGuard();

	constructor(deps: AutosaveDeps) {
		this.#deps = deps;
		this.#lastLoadCount = deps.getLoadCount();
		this.#lastSeedNonce = deps.getSeedNonce?.() ?? 0;
		this.#engine = createAutosaveController({
			getActivation: () => ({ active: deps.getEnabled() }),
			getFilePath: deps.getFilePath,
			getIntervalMs: deps.getIntervalMs,
			serialize: (draft) => this.#serialize(draft),
			...(deps.getInlineEditingId ? { getInlineEditingId: deps.getInlineEditingId } : {}),
			getStorage: () => deps.getStorage?.(),
			onSaved: (bytes) => deps.onSaved?.(bytes),
			writeSnapshot: saveAutosaveSnapshot,
		});
		this.#state = controllerState(this.#engine);

		// The activation, the cadence and the key apply at once.
		$effect(() => {
			deps.getEnabled();
			deps.getIntervalMs();
			deps.getFilePath();
			untrack(() => this.#engine.refresh());
		});

		// Watch edits: reassigning the (immutable) slide array on each edit fires
		// this effect. A load bumps `getLoadCount()` (and the seed nonce), which
		// is a reseed (clear dirty) rather than an edit.
		$effect(() => {
			const loadCount = deps.getLoadCount();
			const seedNonce = deps.getSeedNonce?.() ?? 0;
			// Track the stores so edits (reassignment) re-run the effect.
			deps.getSlides();
			deps.getSlideMasters?.();
			deps.getNotesMaster?.();
			deps.getHandoutMaster?.();
			deps.getSections?.();
			const held = deps.getPreviewHeld?.() ?? false;
			const documentKey = deps.getDocumentKey?.();

			untrack(() => {
				if (!this.#started) {
					this.#started = true;
					this.#lastLoadCount = loadCount;
					this.#lastSeedNonce = seedNonce;
					return;
				}
				// Either half of a load counts as a reseed. They arrive in separate
				// flushes (see `getSeedNonce`), so both have to be able to answer.
				if (loadCount !== this.#lastLoadCount || seedNonce !== this.#lastSeedNonce) {
					this.#lastLoadCount = loadCount;
					this.#lastSeedNonce = seedNonce;
					this.#engine.markClean();
					return;
				}
				if (documentKey && !this.#previewGuard.isEdit(documentKey, held)) {
					return;
				}
				this.#engine.markEdited();
			});
		});
		$effect(() => () => this.#engine.destroy());
	}

	/** Current autosave lifecycle status (reactive). */
	get status(): AutosaveStatus {
		return this.#state.current.status.state;
	}

	/** Whether there are edits no snapshot has captured yet (reactive). */
	get isDirty(): boolean {
		return this.#state.current.dirty;
	}

	/** Epoch ms of the last successful save, or null (reactive). */
	get lastSavedAt(): number | null {
		return this.#state.current.lastSavedAt;
	}

	/**
	 * Serialize the current slides to `.pptx` bytes, or null when unavailable,
	 * through the shared `saveDeckWithPassword` with an explicit
	 * `recoverySnapshotIntent`, so a snapshot is always a plain ZIP that
	 * recovery can reopen.
	 */
	async #serialize(draft: InlineTextDraft | null): Promise<Uint8Array | null> {
		const handler = this.#deps.getHandler();
		if (!handler) {
			return null;
		}
		const slideMasters = this.#deps.getSlideMasters?.();
		const notesMaster = this.#deps.getNotesMaster?.();
		const handoutMaster = this.#deps.getHandoutMaster?.();
		const sections = this.#deps.getSections?.();
		const hasMasters = Boolean(slideMasters?.length || notesMaster || handoutMaster);
		return saveDeckWithPassword(
			handler,
			// Typing not committed yet is part of the snapshot (AutoRecover keeps it).
			[...applyInlineTextDraft(this.#deps.getSlides(), draft)],
			hasMasters
				? {
						slideMasters,
						notesMaster,
						handoutMaster,
						sections: sections?.length ? sections : undefined,
					}
				: { sections: sections?.length ? sections : undefined },
			recoverySnapshotIntent(this.#deps.getSaveIntent?.()),
		);
	}

	/** In-progress inline text: a changed draft is an edit and goes into the snapshot. */
	noteInlineDraft(draft: InlineTextDraft | null): void {
		this.#engine.noteInlineDraft(draft);
	}

	/** Force an immediate save, bypassing the interval. */
	save(): Promise<void> {
		return this.#engine.saveNow();
	}
}
