import { dispatchIsolatedCommand } from 'ooxml-core/docx/ui';
import { closeHistory } from 'prosemirror-history';
import type { Node, ResolvedPos } from 'prosemirror-model';
import { Selection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export function dispatchTableEdit(view: EditorView, transaction: Transaction): void {
	if (!transaction.docChanged) return;
	dispatchIsolatedCommand(view.state, view.dispatch, transaction);
	view.dispatch(closeHistory(view.state.tr));
}

/** Insertion keeps the same textblock and offset when the table is replaced. */
export function tableCaret(
	tr: Transaction,
	table: Node,
	tablePos: number,
	caret: ResolvedPos,
): Selection {
	let position: number | undefined;
	table.descendants((node, offset) => {
		if (node === caret.parent) position ??= tablePos + offset + 2 + caret.parentOffset;
		return position === undefined;
	});
	return Selection.near(tr.doc.resolve(position ?? tablePos + 2));
}
