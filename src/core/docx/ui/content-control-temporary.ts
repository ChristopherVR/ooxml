// A temporary content control (`w:temporary`) is removed, and its content kept as ordinary text,
// once that content is edited (ECMA-376 17.5.2.43). The commands that replace a control's content
// remove it themselves; this covers ordinary typing, deleting and pasting inside one.
import { isHistoryTransaction } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { Mapping } from 'prosemirror-transform';
import { removedByEdit } from '../content-control-edit';
import { CONTROL_EDIT_META } from './content-control-commands';
import { controlOccurrences, type Occurrence, type Span } from './content-control-occurrences';
import { rowControlSpans } from './content-control-rows';
import { controlOfMark, controlSpans, tableControlSpans } from './content-control-spans';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';

/**
 * The control's content as its text, with each object as a character and a picture by its parts,
 * so replacing a picture is an edit too. Formatting is not: it changes no content.
 */
const contentOf = (state: EditorState, range: { from: number; to: number }): string =>
	state.doc.textBetween(range.from, range.to, '\n', (leaf) =>
		leaf.type.name === 'image'
			? `\uFFFC${String(leaf.attrs.partName)}\uFFFC${String(leaf.attrs.svgPartName ?? '')}`
			: '\uFFFC',
	);

/** Takes the control of `span` off its text, paragraphs, table, cell or row; its content stays. */
export function removeControl(tr: Transaction, span: Span): Transaction {
	const { id } = span.control;
	if (span.level === 'row') {
		const row = tr.doc.nodeAt(span.from);
		const { contentControl: _control, ...properties } = row?.attrs.properties ?? {};
		if (row) tr.setNodeMarkup(span.from, undefined, { ...row.attrs, properties });
		return tr;
	}
	if (span.level === 'cell') {
		const cell = tr.doc.nodeAt(span.from);
		if (cell) tr.setNodeMarkup(span.from, undefined, { ...cell.attrs, contentControl: null });
		return tr;
	}
	tr.doc.nodesBetween(span.from, span.to, (node, pos) => {
		if (span.level === 'inline') {
			for (const mark of node.marks)
				if (mark.type.name === 'contentControl' && controlOfMark(mark)?.id === id)
					tr.removeMark(pos, pos + node.nodeSize, mark);
			return true;
		}
		// Its own paragraphs or table: not a cell around them, nor a paragraph in the table, whose
		// control may share its id.
		if (pos < span.from) return true;
		if (BLOCKS.has(node.type.name) && node.attrs.contentControl?.id === id)
			tr.setNodeMarkup(pos, undefined, { ...node.attrs, contentControl: null });
		return false;
	});
	return tr;
}

const BLOCKS = new Set(['paragraph', 'table']);

/**
 * The temporary controls of `doc` an edit would remove. Controls that are not temporary, or that
 * may not be deleted, are left out, so a control that shares an id with one is never taken for it.
 */
function temporaryControls(doc: ProseMirrorNode): Occurrence[] {
	const spans = controlSpans(doc);
	const inline = spans.filter((span) => span.level === 'inline');
	// In document order: the paragraphs and tables of one block control are its spans in turn.
	const candidates = [...spans, ...tableControlSpans(doc), ...rowControlSpans(doc)];
	return controlOccurrences(
		doc,
		candidates.filter((span) => removedByEdit(span.control)),
		inline,
	);
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
	const was = temporaryControls(before.doc);
	if (!was.length) return null;
	const mapping = new Mapping();
	for (const transaction of transactions) mapping.appendMapping(transaction.mapping);
	const now = temporaryControls(after.doc);
	const removed = new Set<Occurrence>();
	// The inline controls the edit removed, or whose text it deleted.
	const dropped = new Set<string>();
	const tr = after.tr;
	for (const old of was) {
		// What the control became: the controls of its level and id where its range went.
		const ends = [mapping.map(old.from, 1), mapping.map(old.to, -1)];
		const [from, to] = [Math.min(...ends), Math.max(...ends)];
		const became = now.filter(
			(occurrence) => occurrence.key === old.key && occurrence.from <= to && occurrence.to >= from,
		);
		const inlineId = old.spans[0]?.level === 'inline' ? old.spans[0].control.id : undefined;
		if (!became.length && inlineId) dropped.add(inlineId);
		// Its whole range: what splits its spans (a nested control, a table) is its content too.
		const content = became.map((occurrence) => contentOf(after, occurrence)).join('\n');
		if (!became.length || content === contentOf(before, old)) continue;
		if (inlineId) dropped.add(inlineId);
		// Every span of an edited control goes: Word removes the whole control.
		for (const occurrence of became) {
			if (removed.has(occurrence)) continue;
			removed.add(occurrence);
			for (const span of occurrence.spans) removeControl(tr, span);
		}
	}
	// Deleting a selection keeps its marks for the text typed next. They stay through the removal,
	// whose steps would clear them, except the mark of a control it took away, which would bring
	// that control back around the typed text.
	const stored = after.storedMarks;
	if (stored) {
		const kept = stored.filter((mark) => {
			const control = mark.type.name === 'contentControl' ? controlOfMark(mark) : undefined;
			return !control || !dropped.has(control.id) || !removedByEdit(control);
		});
		if (tr.docChanged || kept.length !== stored.length) tr.setStoredMarks(kept);
	}
	return tr.docChanged || tr.storedMarksSet ? tr.setMeta(CONTROL_EDIT_META, true) : null;
}
