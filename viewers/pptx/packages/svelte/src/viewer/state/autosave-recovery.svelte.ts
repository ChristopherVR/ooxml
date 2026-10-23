import type {
	AutosaveRecord,
	AutosaveRecoveryController as SharedRecoveryController,
	AutosaveRecoveryPrompt,
	AutosaveRecoveryState,
	DocumentStorage,
} from 'ooxml-ui/pptx';
import {
	createAutosaveRecoveryController,
	discardAutosaveRecovery,
	probeAutosaveRecovery,
} from 'ooxml-ui/pptx';
import { untrack } from 'svelte';

import { controllerState } from './controller-state';
import type { ControllerState } from './controller-state';

/**
 * autosave-recovery.svelte.ts: the Svelte adapter of the shared crash-recovery
 * offer (`createAutosaveRecoveryController` in `ooxml-ui/pptx`).
 *
 * It registers its own probing `$effect` in the constructor, so a caller only
 * has to construct it during component initialisation and render `prompt` when
 * it is non-null. The probe runs once per loaded deck (`getLoadCount`).
 */

export interface AutosaveRecoveryDeps {
	/** IndexedDB record key (host `filePath`). Nothing to look up without one. */
	getFilePath: () => string | undefined;
	/** Public document name rendered in the recovery message. */
	getFileName?: () => string | undefined;
	/** Host `storage` prop: recovery snapshots go to its recovery methods when it has them. */
	getStorage?: () => DocumentStorage | undefined;
	/**
	 * Whether recovery snapshots are permitted at all, which is the HOST prop
	 * (`autosave !== false`) and not the user's toggle: someone who merely
	 * switched AutoSave off should still be offered the pre-crash snapshot that
	 * was written while it was on.
	 */
	getAutosaveAllowed: () => boolean;
	getLoading: () => boolean;
	getError: () => string | null;
	getSlideCount: () => number;
	/** Monotonic load counter: the probe runs once per loaded deck. */
	getLoadCount: () => number;
	/** The loader's own load path, the same one the backstage recent list uses. */
	load: (bytes: Uint8Array) => Promise<void> | void;
	/** The host's `onautosaverecovery`: called with the record when a snapshot is offered. */
	onOffered?: (record: AutosaveRecord) => void;
}

export class AutosaveRecoveryController {
	readonly #controller: SharedRecoveryController;
	readonly #state: ControllerState<AutosaveRecoveryState>;

	constructor(deps: AutosaveRecoveryDeps) {
		this.#controller = createAutosaveRecoveryController({
			getStorage: () => deps.getStorage?.(),
			load: (bytes) => deps.load(bytes),
			onOffered: (record) => deps.onOffered?.(record),
			probe: probeAutosaveRecovery,
			discardRecord: discardAutosaveRecovery,
		});
		this.#state = controllerState(this.#controller);
		$effect(() => {
			const input = {
				loadKey: deps.getLoadCount(),
				filePath: deps.getFilePath(),
				fileName: deps.getFileName?.(),
				loading: deps.getLoading(),
				error: deps.getError(),
				slideCount: deps.getSlideCount(),
				autosaveAllowed: deps.getAutosaveAllowed(),
			};
			untrack(() => this.#controller.check(input));
		});
		$effect(() => () => this.#controller.destroy());
	}

	/** The prompt to render, or null when there is nothing to recover (reactive). */
	get prompt(): AutosaveRecoveryPrompt | null {
		return this.#state.current.prompt;
	}

	/** The destructive delete is still pending (reactive). */
	get discarding(): boolean {
		return this.#state.current.discarding;
	}

	/** Take the snapshot and load its bytes in place without deleting it. */
	restore(): Promise<void> {
		return this.#controller.restore();
	}

	/** Decline the snapshot: delete it, and never load it. */
	discard(): Promise<void> {
		return this.#controller.discard();
	}

	/** Close the prompt without touching the stored snapshot. */
	dismiss(): void {
		this.#controller.dismiss();
	}
}
