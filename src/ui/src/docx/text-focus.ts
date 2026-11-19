import { TextSelection } from 'prosemirror-state';
import type { EditorCore } from './editor-core';

/** A selection in the body text, as document positions (`from` equals `to` for a caret). */
export interface DocxTextRange {
	from: number;
	to: number;
}

export function textSelection(core: EditorCore): DocxTextRange | null {
	const selection = core.view?.state.selection;
	return selection ? { from: selection.from, to: selection.to } : null;
}

/** Focuses the body text, with the selection at `range` clamped to the document when given. */
export function focusText(core: EditorCore, range?: DocxTextRange | null): void {
	const view = core.view;
	if (!view) return;
	if (range) {
		const { doc } = view.state;
		const at = (position: number) => doc.resolve(Math.max(0, Math.min(position, doc.content.size)));
		const selection = TextSelection.between(at(range.from), at(range.to));
		view.dispatch(view.state.tr.setSelection(selection).scrollIntoView());
	}
	view.focus();
}
