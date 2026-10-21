// A temporary content control (`w:temporary`) is removed, and its content kept as ordinary text,
// once that content is edited (ECMA-376 17.5.2.43). The commands that replace a control's content
// remove it themselves; this covers ordinary typing, deleting and pasting inside one.
import { isHistoryTransaction } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { Mapping } from 'prosemirror-transform';
import { removedByEdit } from '../content-control-edit';
import { CONTROL_EDIT_META } from './content-control-commands';
import { rowControlSpans, type RowControlSpan } from './content-control-rows';
import { controlOfMark, controlSpans, type ControlSpan } from './content-control-spans';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';

type Span = ControlSpan | RowControlSpan;

/**
 * The control's content as its text, with each object as a character and a picture by its parts,
 * so replacing a picture is an edit too. Formatting is not: it changes no content.
 */
const contentOf = (state: EditorState, span: Span): string =>
	state.doc.textBetween(span.from, span.to, '\n', (leaf) =>
		leaf.type.name === 'image'
			? `\uFFFC${String(leaf.attrs.partName)}\uFFFC${String(leaf.attrs.svgPartName ?? '')}`
			: '\uFFFC',
	);

/** Takes the control of `span` off its text, paragraphs, cell or row; its content stays as it is. */
export function removeControl(tr: Transaction, span: Span): Transaction {
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

/** One control and its range: its spans, more than one when it holds what splits them. */
interface Occurrence {
	key: string;
	spans: Span[];
	from: number;
	to: number;
}

/**
 * Whether `span` goes on with the control of its level and id that ends at `end`: right after
 * it, or past only what splits one control, other inline controls in its text (a nested control)
 * and, among its paragraphs, blocks other than paragraphs (a table) and paragraphs of another
 * control (a nested block control).
 */
function continues(
	doc: ProseMirrorNode,
	inline: readonly Span[],
	end: number,
	span: Span,
): boolean {
	if (span.from === end) return true;
	if (span.level === 'inline') {
		let at = end;
		for (const other of inline) if (other.from === at && other.to <= span.from) at = other.to;
		return at === span.from;
	}
	if (span.level !== 'block') return false;
	const $end = doc.resolve(end);
	if (doc.resolve(span.from).parent !== $end.parent) return false;
	let split = true;
	$end.parent.forEach((node, offset) => {
		const pos = $end.start() + offset;
		if (pos < end || pos >= span.from || node.type.name !== 'paragraph') return;
		if (!node.attrs.contentControl) split = false;
	});
	return split;
}

/**
 * The temporary controls of `doc` an edit would remove. Controls that are not temporary, or that
 * may not be deleted, are left out, so a control that shares an id with one is never taken for it.
 */
function temporaryControls(doc: ProseMirrorNode): Occurrence[] {
	const spans = controlSpans(doc);
	const inline = spans.filter((span) => span.level === 'inline');
	const found: Occurrence[] = [];
	const last = new Map<string, Occurrence>();
	for (const span of [...spans, ...rowControlSpans(doc)]) {
		if (!removedByEdit(span.control)) continue;
		const key = `${span.level}:${span.control.id}`;
		const previous = last.get(key);
		if (previous && continues(doc, inline, previous.to, span)) {
			previous.spans.push(span);
			previous.to = span.to;
			continue;
		}
		const occurrence = { key, spans: [span], from: span.from, to: span.to };
		found.push(occurrence);
		last.set(key, occurrence);
	}
	return found;
}

const contentOfAll = (state: EditorState, spans: readonly Span[]): string =>
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
	const was = temporaryControls(before.doc);
	if (!was.length) return null;
	const mapping = new Mapping();
	for (const transaction of transactions) mapping.appendMapping(transaction.mapping);
	const now = temporaryControls(after.doc);
	const removed = new Set<Occurrence>();
	const tr = after.tr;
	for (const old of was) {
		// What the control became: the controls of its level and id where its range went.
		const ends = [mapping.map(old.from, 1), mapping.map(old.to, -1)];
		const [from, to] = [Math.min(...ends), Math.max(...ends)];
		const became = now.filter(
			(occurrence) => occurrence.key === old.key && occurrence.from <= to && occurrence.to >= from,
		);
		const content = became.map((occurrence) => contentOfAll(after, occurrence.spans)).join('\n');
		if (!became.length || content === contentOfAll(before, old.spans)) continue;
		// Every span of an edited control goes: Word removes the whole control.
		for (const occurrence of became) {
			if (removed.has(occurrence)) continue;
			removed.add(occurrence);
			for (const span of occurrence.spans) removeControl(tr, span);
		}
	}
	return tr.docChanged ? tr.setMeta(CONTROL_EDIT_META, true) : null;
}
