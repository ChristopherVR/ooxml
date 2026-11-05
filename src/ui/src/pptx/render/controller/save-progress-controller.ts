/**
 * The progress of a Save to the host's storage, shared by every binding.
 *
 * A storage that is slow (a OneDrive upload of a large deck) reports its steps through
 * `DocumentSaveMeta.onProgress`. A binding wraps the save in {@link SaveProgressController.track}
 * and hands the status bar {@link saveProgressStatus} of the snapshot: the visible text updates
 * on every step, the polite announcement only on a phase change or each 10% of a transfer
 * (`saveProgressStepChanged`). When the save settles the state is empty again and the status bar
 * goes back to "Unsaved changes" / "All saved".
 *
 * @module render/controller/save-progress-controller
 */

import type { DocumentSaveMeta, DocumentSaveProgress, DocumentStorage } from 'ooxml-core/embedding';

import {
	SAVE_PROGRESS_TEXT,
	saveProgressLabel,
	saveProgressStepChanged,
} from '../../../files/save-progress-label';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

export interface SaveProgressState {
	/** The latest step; undefined while no save reports one. */
	readonly progress?: DocumentSaveProgress;
	/** The last step that was worth announcing. */
	readonly announced?: DocumentSaveProgress;
}

export interface SaveProgressController extends PptxController<SaveProgressState> {
	/** Record one step of the running save. */
	report(progress: DocumentSaveProgress): void;
	/** Forget the progress (the save settled). */
	clear(): void;
	/** `storage` with its saves reporting here (see {@link withSaveProgress}); one wrapper per storage. */
	wrap<S extends DocumentStorage>(storage: S): S;
	/**
	 * Run a save with progress reporting: `run` receives the `onProgress` callback to put in the
	 * save meta. The progress is cleared when it resolves or rejects; the result is passed on.
	 */
	track<T>(run: (onProgress: (progress: DocumentSaveProgress) => void) => Promise<T>): Promise<T>;
}

export function createSaveProgressController(): SaveProgressController {
	const store = createControllerStore<SaveProgressState>({});
	let runs = 0;
	const report = (progress: DocumentSaveProgress): void => {
		store.update((current) =>
			saveProgressStepChanged(current.announced, progress)
				? { progress, announced: progress }
				: { progress, ...(current.announced ? { announced: current.announced } : {}) },
		);
	};
	const wrapped = new WeakMap<DocumentStorage, DocumentStorage>();
	const controller: SaveProgressController = {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		wrap<S extends DocumentStorage>(storage: S): S {
			let wrapper = wrapped.get(storage);
			if (!wrapper) {
				wrapper = withSaveProgress(storage, controller);
				wrapped.set(storage, wrapper);
			}
			return wrapper as S;
		},
		report,
		clear: () => store.set({}),
		async track(run) {
			const id = ++runs;
			try {
				return await run((progress) => {
					if (id === runs) {
						report(progress);
					}
				});
			} finally {
				if (id === runs) {
					store.set({});
				}
			}
		},
	};
	return controller;
}

/** The translation key of each English progress text, so a dictionary can localise it. */
export const SAVE_PROGRESS_TEXT_KEYS: Readonly<Record<string, string>> = {
	[SAVE_PROGRESS_TEXT.uploadingStart]: 'pptx.autosave.saving',
	[SAVE_PROGRESS_TEXT.uploading]: 'pptx.autosave.progress.uploading',
	[SAVE_PROGRESS_TEXT.checking]: 'pptx.autosave.progress.checking',
	[SAVE_PROGRESS_TEXT.downloading]: 'pptx.autosave.progress.downloading',
	[SAVE_PROGRESS_TEXT.downloadingUnknown]: 'pptx.autosave.progress.downloadingUnknown',
	[SAVE_PROGRESS_TEXT.merging]: 'pptx.autosave.progress.merging',
};

/** What the status bar shows and announces for a save in progress; undefined when none runs. */
export interface SaveProgressStatus {
	/** Shown in the save indicator, updated on every step. */
	readonly text: string;
	/** Placed in a polite live region, updated only on steps worth announcing. */
	readonly announcement: string;
}

/**
 * The status for a snapshot of {@link createSaveProgressController}. `translate` is the binding's
 * `t`; a missing key falls back to the English text.
 */
export function saveProgressStatus(
	state: SaveProgressState,
	translate?: (key: string) => string,
): SaveProgressStatus | undefined {
	if (!state.progress) {
		return undefined;
	}
	const t = (text: string): string => {
		const key = SAVE_PROGRESS_TEXT_KEYS[text];
		const translated = key && translate ? translate(key) : text;
		return !translated || translated === key ? text : translated;
	};
	return {
		text: saveProgressLabel(state.progress, t),
		announcement: saveProgressLabel(state.announced ?? state.progress, t),
	};
}

/**
 * `storage` with its `save` reporting to `controller`: every save the viewer makes through it (File
 * > Save, Save As, a file opened from a cloud place) passes `meta.onProgress`, chained after the
 * caller's own. A binding wraps the storage it hands to its save paths once (and memoises the
 * wrapper per storage); everything else on the storage is reached through the prototype.
 */
export function withSaveProgress<S extends DocumentStorage>(
	storage: S,
	controller: SaveProgressController,
): S {
	return Object.create(storage, {
		save: {
			value: (bytes: Uint8Array, meta: DocumentSaveMeta) =>
				controller.track((report) =>
					storage.save(bytes, {
						...meta,
						onProgress: (progress) => {
							report(progress);
							meta.onProgress?.(progress);
						},
					}),
				),
			enumerable: true,
		},
	}) as S;
}
