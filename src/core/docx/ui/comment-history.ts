import type { Node } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import type { Comment } from '../model';

/** Local history only. Yjs rooms keep independent thread records in their dedicated maps. */
export const commentHistoryAttributes = { commentThreads: { default: null } };

export function commentThreadsFromDoc(doc: Node): Comment[] | undefined {
	if (typeof doc.attrs.commentThreads !== 'string') return;
	try {
		const value: unknown = JSON.parse(doc.attrs.commentThreads);
		if (
			Array.isArray(value) &&
			value.every(
				(entry) =>
					entry &&
					typeof entry.id === 'string' &&
					typeof entry.author === 'string' &&
					typeof entry.text === 'string',
			)
		)
			return value as Comment[];
	} catch {
		/* Preserve the prior model when the local snapshot is unavailable. */
	}
}

export function setCommentThreads(tr: Transaction, comments: readonly Comment[]): boolean {
	if (!commentThreadsFromDoc(tr.doc)) return false;
	const value = JSON.stringify(comments);
	if (value !== tr.doc.attrs.commentThreads) tr.setDocAttribute('commentThreads', value);
	return true;
}

/** Resolution and replies are their own undo events, separate from surrounding text edits. */
export function updateCommentThreads(view: EditorView, comments: readonly Comment[]): boolean {
	if (!view.editable) return false;
	const tr = view.state.tr;
	if (!setCommentThreads(tr, comments)) return false;
	if (tr.docChanged) view.dispatch(closeHistory(tr));
	return true;
}
