import { deliverSavedPresentation, savedPresentationFileName } from 'ooxml-ui/pptx';
import type { DocumentStorage } from 'ooxml-ui/pptx';
import type { PptxSaveFormat } from 'pptx-viewer-core';

import type { EditorState } from './editor-state.svelte';

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
): EditingApi {
	return {
		undo: () => editor.undo(),
		redo: () => editor.redo(),
		canUndo: () => editor.canUndo,
		canRedo: () => editor.canRedo,
		deleteSelected: () => editor.deleteSelected(),
		getSelectedElementId: () => editor.selectedElementId,
		save: (format) => editor.save(format),
		// Every name below goes through the shared save-name decision, so a host
		// that passes the deck it opened (`report.ppt`) gets `report.pptx` back
		// rather than a `.ppt` whose bytes are an OpenXML package. With a host
		// `storage` the bytes go to `storage.save()` instead of a download.
		downloadAs: async (format, fileName) => {
			const bytes = await editor.save(format);
			await deliverSavedPresentation({
				storage: getStorage(),
				bytes,
				format,
				fileName: savedPresentationFileName(fileName, format),
				reason: 'saveAs',
			});
		},
		downloadPptx: async (fileName) => {
			const bytes = await editor.save('pptx');
			await deliverSavedPresentation({
				storage: getStorage(),
				bytes,
				format: 'pptx',
				fileName: savedPresentationFileName(fileName),
				reason: 'save',
			});
		},
	};
}
