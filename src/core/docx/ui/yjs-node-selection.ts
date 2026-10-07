import * as Y from 'yjs';
import { NodeSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	getRelativeSelection,
	initProseMirrorDoc,
	relativePositionToAbsolutePosition,
} from 'y-prosemirror';

/** y-prosemirror 1.3 restores a node selection without checking whether its node was deleted.
 * Convert only that stale selection to a cursor before the binding applies remote content. */
export function guardDeletedNodeSelection(
	doc: Y.Doc,
	view: EditorView,
	binding: Parameters<typeof getRelativeSelection>[0],
): () => void {
	const guard = () => {
		const relative = binding.beforeTransactionSelection;
		const selection = view.state.selection;
		if (relative?.type !== 'node' || !(selection instanceof NodeSelection)) return;
		const current = initProseMirrorDoc(binding.type, view.state.schema);
		const anchor = relative.anchor
			? relativePositionToAbsolutePosition(doc, binding.type, relative.anchor, current.mapping)
			: null;
		const target = anchor === null ? undefined : current.doc.nodeAt(anchor);
		if (target && NodeSelection.isSelectable(target) && target.type === selection.node.type) return;
		binding.beforeTransactionSelection = { ...relative, type: 'text', head: relative.anchor };
	};
	doc.on('beforeObserverCalls', guard);
	return () => doc.off('beforeObserverCalls', guard);
}
