import type { DocumentModel } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { docToModel } from './model-adapter';
import { schema } from './schema';

/** The editor's document model, brought up to date with a collaboration session's shared state. */
export function currentModel(core: EditorCore): DocumentModel {
	if (core.collab.yjs) {
		core.model = docToModel(core.collab.yjs.state(schema).doc, core.model);
		if (core.collab.yjs.sharedComments)
			core.model = { ...core.model, comments: core.collab.yjs.comments.all() };
	}
	return core.model;
}
