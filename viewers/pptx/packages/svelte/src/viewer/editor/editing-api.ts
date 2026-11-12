import {
	createSaveRevisionController,
	deliverSavedPresentation,
	isCleanAfterSave,
	savedPresentationFileName,
	savedPresentationRevision,
} from 'ooxml-ui/pptx';
import type { DocumentStorage, PptxSaveAcknowledgement, PptxSaveSnapshot } from 'ooxml-ui/pptx';
import type { PptxSaveFormat } from 'pptx-viewer-core';

import type { EditorState } from './editor-state.svelte';

/** What a save is built from, each reassigned (never mutated) on an edit; autosave watches the same. */
function saveChangeSources(editor: EditorState): readonly unknown[] {
	return [
		editor.slides,
		editor.templateElementsBySlideId,
		editor.slideMasters,
		editor.notesMaster,
		editor.handoutMaster,
		editor.sections,
	];
}

/** The imperative editing API exposed on the `PowerPointViewer` instance. */
export interface EditingApi {
	undo(): void;
	redo(): void;
	canUndo(): boolean;
	canRedo(): boolean;
	deleteSelected(): void;
	getSelectedElementId(): string | null;
	save(format?: PptxSaveFormat): Promise<Uint8Array>;
	/** Save As `format` (reported to host storage as `saveAs`). */
	downloadAs(format: PptxSaveFormat, fileName?: string): Promise<void>;
	/** File > Save as `.pptx` (reported to host storage as `save`). */
	downloadPptx(fileName?: string): Promise<void>;
	/** The revision of the document as it is now (see `createSaveRevisionController`). */
	getSaveRevision(): number;
	/** Serialise for a host that persists the deck itself; leaves the document dirty. */
	getSaveSnapshot(): Promise<PptxSaveSnapshot>;
	/** The host persisted `revision`: clean only if it is still current. */
	acknowledgeSave(revision: number): PptxSaveAcknowledgement;
}

/**
 * Build the imperative editing API bound to a live `EditorState`. Extracted
 * from `PowerPointViewer.svelte` so the component only re-exports thin,
 * one-line bindings (Svelte requires the component's own `export`s for its
 * instance API, but the bodies can live elsewhere).
 */
export function createEditingApi(
	editor: EditorState,
	getStorage: () => DocumentStorage | undefined = () => undefined,
	/** A Save / Save As landed with no edit while it was in flight: the document is clean. */
	onSavedClean: () => void = () => undefined,
	/** The host acknowledged the current revision: drop the crash-recovery copy. */
	discardRecovery: () => void = () => undefined,
): EditingApi {
	const saveRevision = createSaveRevisionController({
		getSources: () => saveChangeSources(editor),
		serialize: () => editor.serialize('pptx'),
		markClean: () => {
			editor.dirty = false;
			onSavedClean();
		},
		discardRecovery,
	});
	// Serialise without clearing the dirty flag, deliver, and only then decide: a rejected upload
	// keeps the document unsaved, and so does an edit made while it was in flight.
	const deliver = async (
		format: PptxSaveFormat,
		fileName: string,
		reason: 'save' | 'saveAs',
	): Promise<void> => {
		const revision = savedPresentationRevision(saveChangeSources(editor));
		const bytes = await editor.serialize(format);
		await deliverSavedPresentation({ storage: getStorage(), bytes, format, fileName, reason });
		if (isCleanAfterSave(revision, saveChangeSources(editor))) {
			editor.dirty = false;
			onSavedClean();
		}
	};
	return {
		undo: () => editor.undo(),
		redo: () => editor.redo(),
		canUndo: () => editor.canUndo,
		canRedo: () => editor.canRedo,
		deleteSelected: () => editor.deleteSelected(),
		getSelectedElementId: () => editor.selectedElementId,
		save: (format) => editor.save(format),
		getSaveRevision: () => saveRevision.getRevision(),
		getSaveSnapshot: () => saveRevision.snapshot(),
		acknowledgeSave: (revision) => saveRevision.acknowledge(revision),
		// Every name below goes through the shared save-name decision, so a host
		// that passes the deck it opened (`report.ppt`) gets `report.pptx` back
		// rather than a `.ppt` whose bytes are an OpenXML package. With a host
		// `storage` the bytes go to `storage.save()` instead of a download.
		downloadAs: (format, fileName) =>
			deliver(format, savedPresentationFileName(fileName, format), 'saveAs'),
		downloadPptx: (fileName) => deliver('pptx', savedPresentationFileName(fileName), 'save'),
	};
}
