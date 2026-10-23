import type {
	AutosaveActivation,
	AutosaveRecord,
	AutosaveRecoveryController,
	AutosaveRecoveryOffer,
	DeckSaveIntent,
	DocumentStorage,
	ViewerAutosaveStatus,
} from 'ooxml-ui/pptx';
import {
	autosaveStatusKind,
	createAutosaveController as createSharedAutosaveController,
	createAutosaveRecoveryController,
	discardAutosaveRecovery,
	probeAutosaveRecovery,
	recoverySnapshotIntent,
	saveAutosaveSnapshot,
	saveDeckWithPassword,
} from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';

/**
 * The vanilla adapter of the shared recovery autosave engine
 * (`createAutosaveController` in `ooxml-ui/pptx`) and of the shared recovery
 * offer (`createAutosaveRecoveryController`).
 *
 * The engine follows PowerPoint's AutoRecover: while there are unsaved edits a
 * snapshot (always a plain ZIP, written with `recoverySnapshotIntent`) lands
 * one interval after the first edit, and nothing is written while the deck is
 * unchanged. A recovery snapshot **never clears the editor's dirty flag**: it
 * is a crash-safety net, not the user's Save. Whether autosave runs at all is
 * the shared `resolveAutosaveActivation` verdict, read live through
 * {@link AutosaveControllerDeps.getActivation}.
 *
 * This adapter only feeds the engine from the store (local edits set `dirty`;
 * loads and remote collaboration applies do not) and serialises the deck. The
 * recovery offer is probed once per loaded deck and surfaced both to the host
 * callback ({@link AutosaveControllerDeps.onRecovery}) and to the viewer's own
 * prompt ({@link AutosaveControllerDeps.onRecoveryOffer}).
 */
export type AutosaveStatus = ViewerAutosaveStatus;

/** Fallback verdict when a host wires no activation function (tests, embeds). */
const ALWAYS_ACTIVE: AutosaveActivation = { active: true, toggleAvailable: true };

export interface AutosaveControllerDeps {
	store: Store<ViewerState>;
	/** Live core handler used to serialize the current slides to `.pptx` bytes. */
	getHandler: () => PptxHandler | null;
	/** IndexedDB key for the recovery snapshot (usually the open file's name). */
	filePath: string;
	/** Public document name rendered in the recovery message. */
	fileName?: string;
	/** Host `storage` option: recovery snapshots go to its recovery methods when it has them. */
	getStorage?: () => DocumentStorage | undefined;
	/**
	 * The resolved cadence in ms, read EVERY time the timer is armed so a
	 * File > Options > Save change applies without rebuilding the viewer.
	 */
	getIntervalMs: () => number;
	/**
	 * The Protect-Presentation state, forwarded to the shared save decision.
	 * Deliberately inert here: a recovery snapshot is written in the clear
	 * whatever it says.
	 */
	getSaveIntent?: () => DeckSaveIntent;
	/** Surface status transitions (toolbar indicator + optional host callback). */
	onStatus?: (status: AutosaveStatus) => void;
	/** The serialised bytes of each snapshot written, for the host's `onAutosave` hook. */
	onSaved?: (bytes: Uint8Array) => void;
	/** The raw record, for the host's `onAutosaveRecovery` hook. */
	onRecovery?: (record: AutosaveRecord) => void;
	/** The shared prompt descriptor + record, for the viewer's own dialog. */
	onRecoveryOffer?: (offer: AutosaveRecoveryOffer) => void;
	/** The live shared activation verdict; omitted means "always active". */
	getActivation?: () => AutosaveActivation;
	/**
	 * Whether a snapshot may be OFFERED back. This is the host ceiling alone
	 * (`autosave !== false`): a user who merely switched the toggle off should
	 * still be asked about work a crash left behind.
	 */
	isRecoveryAllowed?: () => boolean;
	/** Load restored bytes through the viewer's load path (the recovery offer's restore). */
	loadFile?: (bytes: Uint8Array) => Promise<void> | void;
}

export interface AutosaveController {
	/** Force an immediate snapshot of unsaved edits (no-op while autosave is not active). */
	saveNow(): Promise<void>;
	/** Current status. */
	getStatus(): AutosaveStatus;
	/** Re-read the activation verdict (the toggle or edit-ability changed). */
	refresh(): void;
	/** Whether new edits currently schedule recovery snapshots. */
	isEnabled(): boolean;
	/** The shared recovery offer (restore and dismiss). */
	getRecovery(): AutosaveRecoveryController;
	/** Tear down the timer + store subscription. */
	destroy(): void;
}

export function createAutosaveController(deps: AutosaveControllerDeps): AutosaveController {
	const { store } = deps;
	const activation = (): AutosaveActivation => deps.getActivation?.() ?? ALWAYS_ACTIVE;
	let disposed = false;
	let loadKey = 0;

	const engine = createSharedAutosaveController({
		getActivation: activation,
		getFilePath: () => deps.filePath,
		getIntervalMs: deps.getIntervalMs,
		serialize: async () => {
			const handler = deps.getHandler();
			if (!handler) {
				return null;
			}
			return saveDeckWithPassword(
				handler,
				store.get().slides,
				undefined,
				recoverySnapshotIntent(deps.getSaveIntent?.()),
			);
		},
		...(deps.getStorage ? { getStorage: deps.getStorage } : {}),
		// A Design gallery hover preview is never part of a recovery snapshot.
		isHeld: () => store.get().deckStyleHeld,
		...(deps.onSaved ? { onSaved: deps.onSaved } : {}),
		writeSnapshot: saveAutosaveSnapshot,
	});

	let status: AutosaveStatus = autosaveStatusKind(engine.getSnapshot().status);
	const unsubscribeEngine = engine.subscribe(() => {
		const next = autosaveStatusKind(engine.getSnapshot().status);
		if (next !== status) {
			status = next;
			deps.onStatus?.(next);
		}
	});

	const recovery = createAutosaveRecoveryController({
		...(deps.getStorage ? { getStorage: deps.getStorage } : {}),
		load: (bytes) => deps.loadFile?.(bytes),
		onOffered: (record) => {
			deps.onRecovery?.(record);
			const prompt = recovery.getSnapshot().prompt;
			if (prompt) {
				deps.onRecoveryOffer?.({ prompt, record });
			}
		},
		probe: probeAutosaveRecovery,
		discardRecord: discardAutosaveRecovery,
	});

	const checkRecovery = (state: ViewerState): void => {
		recovery.check({
			loadKey,
			filePath: deps.filePath,
			fileName: deps.fileName,
			loading: state.loading,
			error: state.error,
			slideCount: state.slides.length,
			autosaveAllowed: deps.isRecoveryAllowed?.() ?? true,
		});
	};

	const unsubscribe = store.subscribe((state, previous) => {
		engine.refresh();
		// A fresh load is a fresh deck: it gets its own recovery probe.
		if (state.loading && !previous.loading) {
			loadKey += 1;
		}
		// A hover preview swaps the slides for a moment; neither it nor the restore is an edit.
		if (!state.deckStyleHeld && !previous.deckStyleHeld) {
			engine.observe(state.dirty, [state.slides]);
		}
		if (!state.dirty && previous.dirty && engine.getSnapshot().status.state !== 'saving') {
			// A manual save cleared the dirty flag: reflect it in the indicator.
			engine.markClean({ saved: true });
		}
		checkRecovery(state);
	});

	// Seed the comparison with the state the controller was built on, so only
	// later local edits count. A deck may already be open (a host that
	// re-created it after a load, for instance).
	engine.observe(store.get().dirty, [store.get().slides]);
	checkRecovery(store.get());

	return {
		saveNow: () =>
			engine.getSnapshot().dirty && activation().active ? engine.saveNow() : Promise.resolve(),
		getStatus: () => status,
		refresh: () => engine.refresh(),
		isEnabled: () => activation().active,
		getRecovery: () => recovery,
		destroy() {
			if (disposed) {
				return;
			}
			disposed = true;
			unsubscribe();
			unsubscribeEngine();
			engine.destroy();
			recovery.destroy();
		},
	};
}
