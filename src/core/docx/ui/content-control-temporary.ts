// A temporary content control (`w:temporary`) is removed, and its content kept as ordinary text,
// once that content is edited (ECMA-376 17.5.2.43). The commands that replace a control's content
// remove it themselves; this covers ordinary typing, deleting and pasting inside one.
import { isHistoryTransaction } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { removedByEdit } from '../content-control-edit';
import type { ContentControl } from '../content-control-model';
import { CONTROL_EDIT_META } from './content-control-commands';
import { controlOfMark, controlSpans, type ControlSpan } from './content-control-spans';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';

/** A control and its range, a row control too: the editor keeps it in the row's properties. */
type Wrapped = Omit<ControlSpan, 'level'> & { level: ControlSpan['level'] | 'row' };

/** The table rows wrapped in a control (`w:tbl > w:sdt > w:tr`), each with the row's range. */
function rowSpans(doc: ProseMirrorNode): Wrapped[] {
	const rows: Wrapped[] = [];
	doc.descendants((node, pos) => {
		const properties = node.type.name === 'tableRow' ? node.attrs.properties : undefined;
		const control = properties?.contentControl as ContentControl | undefined;
		if (control) rows.push({ control, level: 'row', from: pos, to: pos + node.nodeSize });
		return !node.isTextblock;
	});
	return rows;
}

/**
 * The control's content as its text, with each object as a character and a picture by its parts,
 * so replacing a picture is an edit too. Formatting is not: it changes no content.
 */
const contentOf = (state: EditorState, span: Wrapped): string =>
	state.doc.textBetween(span.from, span.to, '\n', (leaf) =>
		leaf.type.name === 'image'
			? `\uFFFC${String(leaf.attrs.partName)}\uFFFC${String(leaf.attrs.svgPartName ?? '')}`
			: '\uFFFC',
	);

/** Takes the control of `span` off its text, paragraphs, cell or row; its content stays as it is. */
export function removeControl(tr: Transaction, span: Wrapped): Transaction {
	const { id } = span.control;
	if (span.level === 'row') {
		const row = tr.doc.nodeAt(span.from);
		const { contentControl: _control, ...properties } = row?.attrs.properties ?? {};
		if (row) tr.setNodeMarkup(span.from, undefined, { ...row.attrs, properties });
		return tr;
	}
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

/**
 * The spans of the controls an edit would remove, by level and id: a control whose text a nested
 * control splits, or whose paragraphs a table separates, has several spans and is one control.
 * Controls that are not temporary, or that may not be deleted, are left out, so a control that
 * shares an id with a temporary one is never taken for it.
 */
function temporaryControls(state: EditorState): Map<string, Wrapped[]> {
	const controls = new Map<string, Wrapped[]>();
	for (const span of [...controlSpans(state.doc), ...rowSpans(state.doc)]) {
		if (!removedByEdit(span.control)) continue;
		const key = `${span.level}:${span.control.id}`;
		controls.set(key, [...(controls.get(key) ?? []), span]);
	}
	return controls;
}

const contentOfAll = (state: EditorState, spans: readonly Wrapped[]): string =>
	spans.map((span) => contentOf(state, span)).join('\n');

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
	const was = temporaryControls(before);
	if (!was.size) return null;
	const tr = after.tr;
	for (const [key, spans] of temporaryControls(after)) {
		const old = was.get(key);
		// Every span of an edited control goes: Word removes the whole control.
		if (old && contentOfAll(before, old) !== contentOfAll(after, spans))
			for (const span of spans) removeControl(tr, span);
	}
	return tr.docChanged ? tr.setMeta(CONTROL_EDIT_META, true) : null;
}
