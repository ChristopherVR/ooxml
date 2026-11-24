/** The selected body content as its own document, for Word's "Selection" page range of a PDF. */
import type { DocumentModel } from 'ooxml-core/docx';
import type { EditorState } from 'prosemirror-state';
import { docToModel } from '../model-adapter';

/**
 * `model` with only the blocks the selection covers, cut at the selection's ends (a partly
 * selected paragraph keeps its formatting and only the selected text). Undefined when the
 * selection is empty. Sections, styles, notes, headers and footers stay the document's own.
 */
export function selectionModel(
	state: EditorState,
	model: DocumentModel,
): DocumentModel | undefined {
	const { from, to, empty } = state.selection;
	if (empty) return undefined;
	const doc = state.doc.cut(from, to);
	if (!doc.childCount || (!doc.textContent.trim() && !hasLeaf(doc))) return undefined;
	return docToModel(doc, model);
}

function hasLeaf(doc: EditorState['doc']): boolean {
	let found = false;
	doc.descendants((node) => {
		if (found) return false;
		if (node.isLeaf && !node.isText) found = true;
		return !found;
	});
	return found;
}
