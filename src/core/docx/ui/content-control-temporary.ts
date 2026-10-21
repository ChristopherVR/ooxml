// A temporary content control (`w:temporary`) is removed, and its content kept as ordinary text,
// once that content is edited (ECMA-376 17.5.2.43). The commands that replace a control's content
// remove it themselves; this covers ordinary typing, deleting and pasting inside one.
import { isHistoryTransaction } from 'prosemirror-history';
import type { EditorState, Transaction } from 'prosemirror-state';
import { removedByEdit } from '../content-control-edit';
import { CONTROL_EDIT_META } from './content-control-commands';
import { controlOfMark, controlSpans, type ControlSpan } from './content-control-spans';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';

/**
 * The control's content as its text, with each object as a character and a picture by its parts,
 * so replacing a picture is an edit too. Formatting is not: it changes no content.
 */
const contentOf = (state: EditorState, span: ControlSpan): string =>
	state.doc.textBetween(span.from, span.to, '\n', (leaf) =>
		leaf.type.name === 'image'
			? `\uFFFC${String(leaf.attrs.partName)}\uFFFC${String(leaf.attrs.svgPartName ?? '')}`
			: '\uFFFC',
	);

/** Takes the control of `span` off its text, paragraphs or cell; its content stays as it is. */
export function removeControl(tr: Transaction, span: ControlSpan): Transaction {
	const { id } = span.control;
	tr.doc.nodesBetween(span.from, span.to, (node, pos) => {
		if (span.level === 'inline') {
			for (const mark of node.marks)
				if (mark.type.name === 'contentControl' && controlOfMark(mark)?.id === id)
					tr.removeMark(pos, pos + node.nodeSize, mark);
			return true;
		}
		if (node.attrs.contentControl?.id === id && 'contentControl' in node.attrs)
			tr.setNodeMarkup(pos, undefined, { ...node.attrs, contentControl: null });
		return true;
	});
	return tr;
}

/** The transaction removing every temporary control whose content `transactions` changed. */
export function removeEditedTemporaryControls(
	transactions: readonly Transaction[],
	before: EditorState,
	after: EditorState,
): Transaction | null {
	const local = transactions.filter(
		(tr) =>
			tr.docChanged &&
			!isHistoryTransaction(tr) &&
			!tr.getMeta(REMOTE_TRANSACTION_META) &&
			tr.getMeta('y-sync$') === undefined,
	);
	if (!local.length || local.length !== transactions.filter((tr) => tr.docChanged).length)
		return null;
	const temporary = controlSpans(before.doc).filter((span) => removedByEdit(span.control));
	if (!temporary.length) return null;
	const was = new Map(temporary.map((span) => [span.control.id, contentOf(before, span)]));
	const tr = after.tr;
	for (const span of controlSpans(after.doc)) {
		const content = was.get(span.control.id);
		if (content !== undefined && content !== contentOf(after, span)) removeControl(tr, span);
	}
	return tr.docChanged ? tr.setMeta(CONTROL_EDIT_META, true) : null;
}
