/**
 * Revision-aware acknowledgement of a Save the host persists asynchronously.
 *
 * A host that owns persistence reads the deck with `getSaveSnapshot()`, writes the bytes
 * wherever it likes (possibly slowly, possibly failing) and then calls
 * `acknowledgeSave(revision)` with the revision the snapshot carried. The viewer never reloads
 * and its undo history is untouched:
 *
 *  - the revision is a positive integer naming one observed state of the document (the same
 *    change sources every binding already hands to Save and autosave: the slides array, the
 *    template-element map and whatever else the serializer reads, each reassigned on an edit);
 *  - an acknowledgement for the CURRENT revision leaves the document clean, clears the editor's
 *    dirty flag and discards the crash-recovery copy (when "discard on save" is on), because
 *    that copy only holds what the saved file already has;
 *  - an acknowledgement for an OLDER revision is accepted (it becomes the saved revision) but
 *    the document stays dirty and the recovery copy is kept: it may hold the later edits;
 *  - an acknowledgement older than one already accepted is stale and ignored, and a revision
 *    that was never issued is unknown; neither changes anything;
 *  - a failed persist is simply never acknowledged, so the document stays dirty.
 *
 * Undo back to the saved content is an edit like any other: the revision moves on, so the
 * document is dirty until the next save. (Undoing and redoing between two reads of the revision
 * that lands on the very same state keeps the revision, since the content is identical.)
 *
 * The revision is observed lazily, whenever the host reads it, takes a snapshot or acknowledges,
 * so a binding needs no extra edit hooks, only its `getSources()`.
 *
 * @module render/controller/save-revision
 */

import { autosaveSnapshotMark, autosaveSourcesChanged } from '../autosave-tick';
import type { AutosaveSnapshotMark } from '../autosave-tick';

/** Serialised bytes together with the revision they were taken at. */
export interface PptxSaveSnapshot {
	/** The revision to hand back to `acknowledgeSave` once the bytes are persisted. */
	readonly revision: number;
	/** The presentation as `.pptx` bytes. */
	readonly content: Uint8Array;
}

/** What `acknowledgeSave` did with a revision. */
export type PptxSaveAcknowledgement =
	/** The revision is current: the document is clean now. */
	| { readonly status: 'clean'; readonly revision: number }
	/** Accepted, but edits were made since: the document stays dirty. */
	| {
			readonly status: 'pending';
			readonly revision: number;
			readonly currentRevision: number;
	  }
	/** Older than a revision already acknowledged: ignored. */
	| { readonly status: 'stale'; readonly revision: number; readonly savedRevision: number }
	/** Not a revision this viewer has issued: ignored. */
	| { readonly status: 'unknown'; readonly revision: number };

export interface SaveRevisionControllerDeps {
	/** The values the serializer reads, fresh each call (see `savedPresentationRevision`). */
	getSources(): readonly unknown[];
	/** Serialise the deck WITHOUT clearing the dirty flag or announcing a save. */
	serialize(): Promise<Uint8Array>;
	/** Clear the editor's unsaved-changes state (and the autosave engine's). */
	markClean(): void;
	/** Drop the crash-recovery copy of the saved document (best effort, honours the option). */
	discardRecovery(): void;
}

export interface SaveRevisionController {
	/** The revision of the document as it is now; changes after every edit that was observed. */
	getRevision(): number;
	/** The newest revision the host acknowledged, or 0 before any. */
	getSavedRevision(): number;
	/** Serialise the deck and tag the bytes with the revision read just before. */
	snapshot(): Promise<PptxSaveSnapshot>;
	/** Tell the viewer the host persisted `revision` (see the module note). */
	acknowledge(revision: number): PptxSaveAcknowledgement;
}

/** Build the per-viewer save revision tracker (see the module note). */
export function createSaveRevisionController(
	deps: SaveRevisionControllerDeps,
): SaveRevisionController {
	let revision = 0;
	let mark: AutosaveSnapshotMark | undefined;
	let savedRevision = 0;

	const observe = (): number => {
		const sources = deps.getSources();
		// A binding that offers no sources cannot be told apart from itself: every read is a new
		// revision, so nothing but an immediate acknowledgement is ever "current" (never falsely clean).
		if (sources.length === 0 || autosaveSourcesChanged(mark, sources)) {
			revision += 1;
			mark = autosaveSnapshotMark('', sources);
		}
		return revision;
	};

	return {
		getRevision: observe,
		getSavedRevision: () => savedRevision,
		async snapshot() {
			// Read before serialising: an edit made while the bytes are produced is newer.
			const at = observe();
			const content = await deps.serialize();
			return { revision: at, content };
		},
		acknowledge(acknowledged) {
			const current = observe();
			if (!Number.isInteger(acknowledged) || acknowledged < 1 || acknowledged > current) {
				return { status: 'unknown', revision: acknowledged };
			}
			if (acknowledged < savedRevision) {
				return { status: 'stale', revision: acknowledged, savedRevision };
			}
			const repeated = acknowledged === savedRevision;
			savedRevision = acknowledged;
			if (acknowledged !== current) {
				return { status: 'pending', revision: acknowledged, currentRevision: current };
			}
			deps.markClean();
			if (!repeated) {
				deps.discardRecovery();
			}
			return { status: 'clean', revision: acknowledged };
		},
	};
}
