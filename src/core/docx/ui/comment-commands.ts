import type { Comment } from '../model';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { TextSelection } from 'prosemirror-state';
import { commentIdsFromNode } from './comment-anchors';
import { hasInlineRunAttributes, setInlineRunFormatting } from './inline-formatting';
import { inlineNodeRun } from './run-adapter';

let commentSerial = 0;
function nextCommentId(idGenerator?: (kind: string) => string): string {
	return idGenerator
		? idGenerator('comment')
		: `dve-comment-${Date.now().toString(36)}-${++commentSerial}`;
}

/** Anchors a new comment to selected text or canonical inline run properties. */
export function addComment(
	view: EditorView,
	author: string,
	text: string,
	idGenerator?: (kind: string) => string,
	inlineElements = true,
): Comment | null {
	if (!view.editable || !view.state.schema.marks.comment) return null;
	const { from, to } = view.state.selection;
	if (from === to) return null;
	const id = nextCommentId(idGenerator);
	let tr = view.state.tr;
	view.state.doc.nodesBetween(from, to, (node, pos) => {
		if (!inlineElements && !node.isText) return;
		if (hasInlineRunAttributes(node)) {
			if (!inlineElements) return;
			const run = inlineNodeRun(node);
			if (run)
				setInlineRunFormatting(tr, pos, {
					...run,
					commentIds: [...new Set([...(run.commentIds ?? []), id])].sort(),
				});
			return;
		}
		if (!node.isText && node.type.name !== 'hardBreak') return;
		const start = Math.max(pos, from);
		const end = Math.min(pos + node.nodeSize, to);
		if (end <= start) return;
		tr = tr.addMark(start, end, view.state.schema.marks.comment!.create({ ids: [id] }));
	});
	if (!tr.docChanged) return null;
	view.dispatch(closeHistory(tr));
	return { id, author, text, resolved: false };
}

/** Removes one comment id from every anchor range that carries it. */
export function removeCommentAnchor(view: EditorView, id: string): void {
	if (!view.editable || !view.state.schema.marks.comment) return;
	let tr = view.state.tr;
	view.state.doc.descendants((node, pos) => {
		if (hasInlineRunAttributes(node)) {
			const run = inlineNodeRun(node);
			if (run?.commentIds?.includes(id)) {
				const remaining = run.commentIds.filter((existing) => existing !== id);
				delete run.commentIds;
				if (remaining.length) run.commentIds = remaining;
				setInlineRunFormatting(tr, pos, run);
			}
		}
		if (!node.isInline) return;
		for (const mark of node.marks) {
			if (mark.type.name !== 'comment' || !(mark.attrs.ids as string[]).includes(id)) continue;
			const remaining = (mark.attrs.ids as string[]).filter((existing) => existing !== id);
			tr = tr.removeMark(pos, pos + node.nodeSize, mark);
			if (remaining.length)
				tr = tr.addMark(pos, pos + node.nodeSize, mark.type.create({ ids: remaining }));
		}
	});
	if (tr.docChanged) view.dispatch(closeHistory(tr));
}

export function deleteComment(comments: Comment[], id: string): Comment[] {
	return comments.filter((comment) => comment.id !== id && comment.parentId !== id);
}
export function resolveComment(comments: Comment[], id: string, resolved: boolean): Comment[] {
	return comments.map((comment) => (comment.id === id ? { ...comment, resolved } : comment));
}
export function replyToComment(
	comments: Comment[],
	parentId: string,
	author: string,
	text: string,
	idGenerator?: (kind: string) => string,
): Comment[] {
	const id = nextCommentId(idGenerator);
	return [...comments, { id, author, text, parentId }];
}

/** Comment ids anchored at the selection, in stable order (replies are not anchored). */
export function commentIdsAtSelection(view: EditorView): string[] {
	const { from, to, empty } = view.state.selection;
	const ids: string[] = [];
	view.state.doc.nodesBetween(
		empty ? Math.max(0, from - 1) : from,
		empty ? from + 1 : to,
		(node) => {
			for (const id of commentIdsFromNode(node)) if (!ids.includes(id)) ids.push(id);
		},
	);
	return ids;
}

/** Where each comment's anchored text starts, in document order. */
export function commentAnchors(view: EditorView): Array<{ id: string; from: number }> {
	const seen = new Set<string>();
	const anchors: Array<{ id: string; from: number }> = [];
	view.state.doc.descendants((node, pos) => {
		for (const id of commentIdsFromNode(node))
			if (!seen.has(id)) {
				seen.add(id);
				anchors.push({ id, from: pos });
			}
	});
	return anchors;
}

/**
 * Review > Previous / Next Comment: moves the caret to the start of the neighbouring commented
 * text, wrapping around. Returns the comment id it landed on, or null when there are none.
 */
export function goToComment(view: EditorView, direction: 'previous' | 'next'): string | null {
	const anchors = commentAnchors(view);
	if (!anchors.length) return null;
	const here = view.state.selection.from;
	const target =
		direction === 'next'
			? (anchors.find((anchor) => anchor.from > here) ?? anchors[0]!)
			: ([...anchors].reverse().find((anchor) => anchor.from < here) ?? anchors.at(-1)!);
	view.dispatch(
		view.state.tr.setSelection(TextSelection.create(view.state.doc, target.from)).scrollIntoView(),
	);
	view.focus();
	return target.id;
}
