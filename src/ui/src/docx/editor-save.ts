import { saveDocx, type DocumentModel } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { modifiedStamp } from './editor-attributes';

/** Saves a projection with the template's original package and any pictures staged in the editor. */
export function saveEditorModel(core: EditorCore, model: DocumentModel): Promise<Uint8Array> {
	const pending = new Map([...core.inserts.pendingMedia, ...(core.collab.yjs?.media.all() ?? [])]);
	const media = pending.size ? pending : undefined;
	if (core.loaded) {
		const options = { modified: modifiedStamp(core.reviewAuthor) };
		// A converted Word 97-2003 document uses its DOCX writer.
		return core.loaded.saveDocx
			? core.loaded.saveDocx(model, media, options)
			: core.loaded.save(model, media, options);
	}
	return saveDocx(model, media);
}
