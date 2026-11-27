import { Fragment, Slice } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { OfficeAddInSelection } from '../ribbon/add-in-runtime-settings';
import { ADD_IN_TEXT_LIMIT } from '../ribbon/add-in-protocol';

/**
 * Paragraph attributes that name one paragraph or record its history: a paragraph an add-in
 * inserts copies the formatting of the one it lands in, never its identity, revision marks,
 * list label, bookmarks, permissions, content control or source property snapshots.
 */
const PARAGRAPH_IDENTITY_ATTRS: Readonly<Record<string, unknown>> = {
	id: '',
	markRevision: null,
	formatRevision: null,
	restoredParagraphPropertiesXml: null,
	sourceParagraphPropertiesXml: null,
	listLabelText: null,
	bookmarks: [],
	permissions: null,
	contentControl: null,
};

function insertedParagraphAttrs(attrs: Readonly<Record<string, unknown>>) {
	const next: Record<string, unknown> = { ...attrs };
	for (const [name, value] of Object.entries(PARAGRAPH_IDENTITY_ATTRS))
		if (name in next) next[name] = value;
	return next;
}

/** Resolves the editor's current view on every request, so a pane opened before it exists works. */
export function wordAddInSelection(current: () => EditorView | undefined): OfficeAddInSelection {
	const live = () => {
		const view = current();
		return view && !view.isDestroyed ? view : undefined;
	};
	return {
		readText() {
			const view = live();
			if (!view) return;
			const { from, to } = view.state.selection;
			if (to - from <= ADD_IN_TEXT_LIMIT) return view.state.doc.textBetween(from, to, '\n');
		},
		writeText(text) {
			const view = live();
			if (!view?.editable) return false;
			const { state } = view;
			const lines = text.split(/\r\n?|\n/);
			const { schema } = state;
			const paragraph = schema.nodes.paragraph;
			const { $from } = state.selection;
			const attrs = insertedParagraphAttrs($from.parent.attrs);
			const marks = state.storedMarks ?? $from.marks();
			const transaction =
				lines.length > 1 && paragraph
					? state.tr.replaceSelection(
							new Slice(
								Fragment.from(
									lines.map((line) =>
										paragraph.create(attrs, line ? schema.text(line, marks) : undefined),
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
