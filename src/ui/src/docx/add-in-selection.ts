import type { EditorView } from 'prosemirror-view';
import type { OfficeAddInSelection } from '../ribbon/add-in-runtime-settings';
import { ADD_IN_TEXT_LIMIT } from '../ribbon/add-in-protocol';

export function wordAddInSelection(current: () => EditorView | undefined): OfficeAddInSelection {
	const view = current();
	const live = () => view && !view.isDestroyed && current() === view && view.editable;
	return {
		readText() {
			if (!view || !live()) return;
			const { from, to } = view.state.selection;
			if (to - from <= ADD_IN_TEXT_LIMIT) return view.state.doc.textBetween(from, to, '\n');
		},
		writeText(text) {
			if (!view || !live()) return false;
			const { state } = view;
			const transaction = state.tr.insertText(text);
			if (!transaction.docChanged) return true;
			view.dispatch(transaction.scrollIntoView());
			return view.state.doc !== state.doc;
		},
	};
}
