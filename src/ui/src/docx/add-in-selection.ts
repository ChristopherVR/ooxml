import { Fragment, Slice } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { OfficeAddInSelection } from '../ribbon/add-in-runtime-settings';
import { ADD_IN_TEXT_LIMIT } from '../ribbon/add-in-protocol';

export function wordAddInSelection(current: () => EditorView | undefined): OfficeAddInSelection {
	const view = current();
	const live = () => view && !view.isDestroyed && current() === view;
	return {
		readText() {
			if (!view || !live()) return;
			const { from, to } = view.state.selection;
			if (to - from <= ADD_IN_TEXT_LIMIT) return view.state.doc.textBetween(from, to, '\n');
		},
		writeText(text) {
			if (!view || !live() || !view.editable) return false;
			const { state } = view;
			const lines = text.split(/\r\n?|\n/);
			const { schema } = state;
			const paragraph = schema.nodes.paragraph;
			const marks = state.storedMarks ?? state.selection.$from.marks();
			const transaction =
				lines.length > 1 && paragraph
					? state.tr.replaceSelection(
							new Slice(
								Fragment.from(
									lines.map((line) =>
										paragraph.create(null, line ? schema.text(line, marks) : undefined),
									),
								),
								1,
								1,
							),
						)
					: state.tr.insertText(text);
			if (!transaction.docChanged) return true;
			view.dispatch(transaction.scrollIntoView());
			return view.state.doc !== state.doc;
		},
	};
}
