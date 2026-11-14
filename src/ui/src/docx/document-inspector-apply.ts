import type { EditorState, Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import {
	removeDocumentInformation,
	type DocumentModel,
	type InspectionCategory,
} from 'ooxml-core/docx';
import { forcesTrackChanges, trackChangesPluginKey } from 'ooxml-core/docx/ui';
import { modelToDoc } from './model-adapter';
import { documentInformationJson } from './document-information-attr';
import type { DocumentInspectorHost } from './document-inspector-dialog';

/**
 * Inspect Document's Remove All as one undoable transaction: the changed body blocks are
 * replaced and every document attribute the removal changed (headers and footers, comments,
 * notes, properties) is set, so Undo brings all of it back. Null when nothing changes.
 */
export function documentInformationRemoval(
	state: EditorState,
	model: DocumentModel,
	categories: readonly InspectionCategory[],
): Transaction | null {
	const next = removeDocumentInformation(model, categories);
	if (next === model) return null;
	const doc = modelToDoc(next);
	const current = state.doc;
	const tr = closeHistory(state.tr);
	const start = current.content.findDiffStart(doc.content);
	if (start !== null) {
		const end = current.content.findDiffEnd(doc.content)!;
		// Overlapping ends (a repeated node) are clamped the way ProseMirror's own diffing does.
		const overlap = Math.max(0, start - Math.min(end.a, end.b));
		tr.replace(start, end.a + overlap, doc.slice(start, end.b + overlap));
	}
	for (const [key, value] of Object.entries(doc.attrs))
		if (current.attrs[key] !== value) tr.setDocAttribute(key, value);
	if (!tr.docChanged) return null;
	// The removal is never itself a tracked change, even with Track Changes on.
	return tr.setMeta(trackChangesPluginKey, { tracked: true });
}

/** Why a category cannot be removed now, or undefined when it can. */
export function removalBlocked(
	model: DocumentModel,
	category: InspectionCategory,
	editable: boolean,
): 'readOnly' | 'protected' | undefined {
	if (!editable) return 'readOnly';
	// Tracked-changes protection keeps every change and comment, as Review > Accept does.
	if (category === 'comments' && forcesTrackChanges(model.protection)) return 'protected';
	return undefined;
}

/** The Document Inspector's Remove All, bound to the editor. */
export function inspectorActions(editor: {
	view(): EditorView | undefined;
	model(): DocumentModel;
	/** Read-only, collaborating or story-locking protection rules out every removal. */
	canEdit(): boolean;
	reportError(cause: unknown): void;
}): Omit<DocumentInspectorHost, 'model' | 'restoreFocus'> {
	const editable = () => Boolean(editor.view()?.editable) && editor.canEdit();
	return {
		blocked: (category) => removalBlocked(editor.model(), category, editable()),
		remove: (category) => {
			const view = editor.view();
			if (!view || removalBlocked(editor.model(), category, editable())) return false;
			try {
				const tr = documentInformationRemoval(view.state, editor.model(), [category]);
				if (tr) view.dispatch(tr);
				return true;
			} catch (cause) {
				editor.reportError(cause);
				return false;
			}
		},
	};
}

/** Records a File > Info property edit (made on the model) on the editor document, off the undo stack. */
export function syncDocumentInformation(view: EditorView | undefined, model: DocumentModel): void {
	if (!view) return;
	const json = documentInformationJson(model);
	if (view.state.doc.attrs.documentInformation === json) return;
	view.dispatch(
		view.state.tr.setDocAttribute('documentInformation', json).setMeta('addToHistory', false),
	);
}
