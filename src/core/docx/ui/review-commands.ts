import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { dispatchIsolatedCommand } from './command-history';
import { moveName } from './review-schema';
import { trackChangesPluginKey } from './track-changes-mode';
import { formattingRevision, resolveFormattingRange } from './review-formatting';
import {
	paragraphFormattingRevision,
	resolveParagraphFormatting,
} from './review-paragraph-formatting';

export interface RevisionRange {
	id: string;
	kind: 'insert' | 'delete' | 'formatChange' | 'paragraphChange';
	/** Node position for a paragraph-format change, mapped before applying a command. */
	paragraphPos?: number;
	from: number;
	to: number;
	author: string;
	/** The move this range is one side of; both sides are accepted or rejected together. */
	move?: string;
}

/** Merges adjacent text/hardBreak nodes sharing the same revision mark id into one range. */
export function collectRevisionRanges(doc: import('prosemirror-model').Node): RevisionRange[] {
	const ranges: RevisionRange[] = [];
	doc.descendants((node, pos) => {
		const paragraph = paragraphFormattingRevision(node);
		if (paragraph)
			ranges.push({
				id: paragraph.id,
				kind: 'paragraphChange',
				author: paragraph.author,
				from: pos + 1,
				to: pos + node.nodeSize - 1,
				paragraphPos: pos,
			});
		if (!node.isText && node.type.name !== 'hardBreak') return;
		const mark = node.marks.find(
			(item) => item.type.name === 'insertion' || item.type.name === 'deletion',
		);
		const format = formattingRevision(node);
		if (mark && format)
			ranges.push({
				id: format.id,
				kind: 'formatChange',
				author: format.author,
				from: pos,
				to: pos + node.nodeSize,
			});
		if (!mark && !format) return;
		const kind = mark ? (mark.type.name === 'insertion' ? 'insert' : 'delete') : 'formatChange';
		const id = String(mark?.attrs.id ?? format!.id);
		const move = mark && moveName(mark.attrs.move);
		const author = String(mark?.attrs.author ?? format!.author);
		const last = ranges.at(-1);
		if (
			last &&
			last.id === id &&
			last.kind === kind &&
			last.author === author &&
			last.move === move &&
			last.to === pos
		)
			last.to = pos + node.nodeSize;
		else
			ranges.push({
				id,
				kind,
				from: pos,
				to: pos + node.nodeSize,
				author,
				...(move ? { move } : {}),
			});
	});
	return ranges;
}

function revisionNear(view: EditorView): RevisionRange | undefined {
	const ranges = collectRevisionRanges(view.state.doc);
	const { from, to } = view.state.selection;
	return (
		ranges.find((range) => range.from < to && range.to > from) ?? ranges.find((r) => r.from >= to)
	);
}

export function hasAnyChange(view: EditorView | undefined): boolean {
	return Boolean(view && collectRevisionRanges(view.state.doc).length);
}
export function hasChangeAtCursor(view: EditorView | undefined): boolean {
	return Boolean(view && revisionNear(view));
}

/**
 * The current ranges `range` resolves with: every range of the same revision id (re-located in
 * case the document changed since collection) and, for a move, every range on either side of it.
 */
function linkedRanges(view: EditorView, range: RevisionRange): RevisionRange[] {
	const current = collectRevisionRanges(view.state.doc);
	const linked = current.filter(
		(item) =>
			item.author === range.author &&
			((item.id === range.id && item.kind === range.kind) ||
				(range.move !== undefined && item.move === range.move)),
	);
	return linked;
}

/** Applies accept or reject to ranges in one transaction, mapping positions as text is removed. */
function resolveRanges(view: EditorView, ranges: RevisionRange[], mode: 'accept' | 'reject') {
	let tr = view.state.tr;
	for (const range of ranges) {
		const from = tr.mapping.map(range.from);
		const to = tr.mapping.map(range.to);
		if (range.kind === 'paragraphChange' && range.paragraphPos !== undefined) {
			resolveParagraphFormatting(tr, tr.mapping.map(range.paragraphPos), mode);
			continue;
		}
		if (range.kind === 'formatChange') {
			resolveFormattingRange(tr, from, to, mode);
			continue;
		}
		const removeText = (mode === 'accept') === (range.kind === 'delete');
		if (removeText) tr = tr.delete(from, to);
		else
			tr = tr.removeMark(
				from,
				to,
				view.state.schema.marks[range.kind === 'insert' ? 'insertion' : 'deletion'],
			);
	}
	return tr.setMeta(trackChangesPluginKey, { tracked: true });
}

/** Accepting an insertion keeps the text; accepting a deletion removes it. Moves resolve both sides. */
export function acceptRevisionRange(view: EditorView, range: RevisionRange): void {
	if (!view.editable) return;
	const tr = resolveRanges(view, linkedRanges(view, range), 'accept');
	dispatchIsolatedCommand(
		view.state,
		(transaction) => view.dispatch(transaction),
		tr.scrollIntoView(),
	);
}
/** Rejecting an insertion removes the text; rejecting a deletion restores it. Moves resolve both sides. */
export function rejectRevisionRange(view: EditorView, range: RevisionRange): void {
	if (!view.editable) return;
	const tr = resolveRanges(view, linkedRanges(view, range), 'reject');
	dispatchIsolatedCommand(
		view.state,
		(transaction) => view.dispatch(transaction),
		tr.scrollIntoView(),
	);
}

export function acceptChangeAtCursor(view: EditorView): boolean {
	if (!view.editable) return false;
	const range = revisionNear(view);
	if (!range) return false;
	acceptRevisionRange(view, range);
	return true;
}
export function rejectChangeAtCursor(view: EditorView): boolean {
	if (!view.editable) return false;
	const range = revisionNear(view);
	if (!range) return false;
	rejectRevisionRange(view, range);
	return true;
}

function applyToAll(view: EditorView, mode: 'accept' | 'reject'): boolean {
	if (!view.editable) return false;
	const ranges = collectRevisionRanges(view.state.doc);
	if (!ranges.length) return false;
	dispatchIsolatedCommand(
		view.state,
		(transaction) => view.dispatch(transaction),
		resolveRanges(view, ranges, mode),
	);
	return true;
}
export const acceptAllChanges = (view: EditorView): boolean => applyToAll(view, 'accept');
export const rejectAllChanges = (view: EditorView): boolean => applyToAll(view, 'reject');

function navigate(view: EditorView, direction: 'next' | 'previous'): boolean {
	const ranges = collectRevisionRanges(view.state.doc);
	if (!ranges.length) return false;
	const { from, to } = view.state.selection;
	const target =
		direction === 'next'
			? (ranges.find((range) => range.from > to) ?? ranges[0])
			: ([...ranges].reverse().find((range) => range.to < from) ?? ranges.at(-1)!);
	if (!target) return false;
	view.dispatch(
		view.state.tr
			.setSelection(TextSelection.create(view.state.doc, target.from, target.to))
			.scrollIntoView(),
	);
	view.focus();
	return true;
}
export const goToNextChange = (view: EditorView): boolean => navigate(view, 'next');
export const goToPreviousChange = (view: EditorView): boolean => navigate(view, 'previous');
