/**
 * The crash-recovery offer every PowerPoint binding adapts.
 *
 * After a deck loads, the controller asks the recovery store whether a
 * snapshot someone else left behind exists for it (`probeAutosaveRecovery`:
 * fresh enough, not written by this page, not already loaded by this tab) and
 * holds the prompt descriptor the binding renders as its own native dialog.
 * Restore loads the bytes through the binding's own load path and keeps the
 * snapshot (it stays recoverable until explicitly discarded); Discard deletes
 * it, and leaves the prompt open when the delete fails so it can be retried.
 *
 * The probe runs once per loaded deck. Bindings that load more than one deck
 * in a viewer pass a `loadKey` that changes with each load; without one the
 * probe runs once per controller. Whether the prompt may be on screen right
 * now (not during a slide show) stays the pure `shouldShowAutosaveRecoveryPrompt`.
 *
 * @module render/controller/autosave-recovery-controller
 */

import type { DocumentStorage } from 'ooxml-core/embedding';

import {
	acceptAutosaveRecovery,
	discardAutosaveRecovery,
	probeAutosaveRecovery,
	shouldProbeAutosaveRecovery,
} from '../autosave-recovery';
import type { AutosaveRecoveryOffer, AutosaveRecoveryPrompt } from '../autosave-recovery';
import {
	acknowledgeAutosaveRecovery,
	clearAutosaveRecoveryAcknowledgement,
} from '../autosave-recovery-acknowledgement';
import type { AutosaveRecord } from '../autosave-store';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** What the recovery dialog reads. */
export interface AutosaveRecoveryState {
	/** What the dialog should say, or null when there is nothing to offer. */
	readonly prompt: AutosaveRecoveryPrompt | null;
	/** The destructive delete is still pending. */
	readonly discarding: boolean;
}

/** The binding's live view of the loaded deck, passed to each {@link AutosaveRecoveryController.check}. */
export interface AutosaveRecoveryCheckInput {
	readonly filePath: string | undefined;
	/** Public document name for the message; the key stays storage-only. */
	readonly fileName?: string | undefined;
	readonly loading: boolean;
	readonly error: string | null | undefined;
	readonly slideCount: number;
	/** The host ceiling alone (`autosave !== false`), not the user's toggle. */
	readonly autosaveAllowed: boolean;
	/** Changes with each loaded deck; the probe runs once per key. */
	readonly loadKey?: unknown;
}

export interface AutosaveRecoveryControllerDeps {
	/** Host `storage`: the snapshot is read from its recovery methods when it has them. */
	getStorage?(): DocumentStorage | undefined;
	/** Load the recovered bytes through the binding's own load path. */
	load(bytes: Uint8Array): Promise<void> | void;
	/** Called with the stored record when a snapshot is offered (the host's recovery event). */
	onOffered?(record: AutosaveRecord): void;
	now?(): number;
	/** Seams for tests; default to the shared store functions. */
	probe?(
		filePath: string,
		now: number,
		displayName: string | undefined,
		storage: DocumentStorage | undefined,
	): Promise<AutosaveRecoveryOffer | null>;
	discardRecord?(record: AutosaveRecord, storage: DocumentStorage | undefined): Promise<void>;
}

export interface AutosaveRecoveryController extends PptxController<AutosaveRecoveryState> {
	/** Probe for a snapshot when the deck is ready and this load was not probed yet. */
	check(input: AutosaveRecoveryCheckInput): void;
	/** Accept: close the prompt and load the snapshot (it stays stored). */
	restore(): Promise<void>;
	/** Decline: delete the snapshot, then close (stays open when the delete fails). */
	discard(): Promise<void>;
	/** Close the prompt without touching the stored snapshot. */
	dismiss(): void;
	destroy(): void;
}

const NEVER = Symbol('never-checked');

/** Build the recovery offer controller (see the module note). */
export function createAutosaveRecoveryController(
	deps: AutosaveRecoveryControllerDeps,
): AutosaveRecoveryController {
	const store = createControllerStore<AutosaveRecoveryState>({ prompt: null, discarding: false });
	const probe = deps.probe ?? probeAutosaveRecovery;
	const discardRecord = deps.discardRecord ?? discardAutosaveRecovery;
	let record: AutosaveRecord | null = null;
	let checkedKey: unknown = NEVER;
	let probeSeq = 0;
	let restoring = false;
	let actionPending = false;
	let destroyed = false;

	const set = (next: Partial<AutosaveRecoveryState>): void => {
		const current = store.getSnapshot();
		if (
			(next.prompt === undefined || next.prompt === current.prompt) &&
			(next.discarding === undefined || next.discarding === current.discarding)
		) {
			return;
		}
		store.set({ ...current, ...next });
	};

	const dismiss = (): void => {
		record = null;
		set({ prompt: null });
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		check(input) {
			if (destroyed) {
				return;
			}
			if (restoring) {
				// The restore's own load is not a deck to offer the snapshot for again.
				checkedKey = input.loadKey;
				return;
			}
			const filePath = input.filePath;
			if (
				!filePath ||
				!shouldProbeAutosaveRecovery({
					alreadyChecked: checkedKey !== NEVER && Object.is(checkedKey, input.loadKey),
					filePath,
					loading: input.loading,
					error: input.error,
					slideCount: input.slideCount,
					autosaveAllowed: input.autosaveAllowed,
				})
			) {
				return;
			}
			checkedKey = input.loadKey;
			const seq = ++probeSeq;
			void probe(filePath, (deps.now ?? Date.now)(), input.fileName, deps.getStorage?.()).then(
				(offer) => {
					// A newer load probed while this one was answering: its answer owns the prompt.
					if (destroyed || seq !== probeSeq || !offer) {
						return;
					}
					record = offer.record;
					set({ prompt: offer.prompt });
					deps.onOffered?.(offer.record);
				},
			);
		},
		async restore() {
			if (actionPending) {
				return;
			}
			const found = record;
			dismiss();
			if (!found) {
				return;
			}
			actionPending = true;
			restoring = true;
			acknowledgeAutosaveRecovery(found);
			try {
				await deps.load(acceptAutosaveRecovery(found));
			} catch (error) {
				clearAutosaveRecoveryAcknowledgement(found);
				throw error;
			} finally {
				restoring = false;
				actionPending = false;
			}
		},
		async discard() {
			if (actionPending) {
				return;
			}
			const found = record;
			if (!found) {
				dismiss();
				return;
			}
			actionPending = true;
			set({ discarding: true });
			try {
				await discardRecord(found, deps.getStorage?.());
				if (record === found) {
					dismiss();
				}
			} catch {
				// Leave the prompt open so a failed discard can be retried.
			} finally {
				actionPending = false;
				set({ discarding: false });
			}
		},
		dismiss,
		destroy() {
			destroyed = true;
		},
	};
}
