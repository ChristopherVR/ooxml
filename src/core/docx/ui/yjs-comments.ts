import * as Y from 'yjs';
import type { CollabSession } from '../../collab/index';
import type { Comment } from '../model';
import type { EditorView } from 'prosemirror-view';
import { ySyncPluginKey } from 'y-prosemirror';
import { addComment } from './comment-commands';

/** Independent records avoid replacing a whole thread when two authors reply offline.
 * Deletion hides concurrent replies until the root is explicitly restored by undo. */
export class WordYjsComments {
	readonly records: Y.Map<Comment>;
	readonly resolved: Y.Map<boolean>;
	readonly deleted: Y.Map<boolean>;
	constructor(
		private readonly session: CollabSession,
		private readonly writable: () => boolean,
		private readonly stopCapturing: () => void,
		private readonly fragment: Y.XmlFragment,
	) {
		this.records = session.doc.getMap('docx:comments');
		this.resolved = session.doc.getMap('docx:comments:resolved');
		this.deleted = session.doc.getMap('docx:comments:deleted');
	}

	all(): Comment[] {
		const visible = (comment: Comment): boolean => {
			if (this.deleted.get(comment.id)) return false;
			if (!comment.parentId) return true;
			const root = this.records.get(comment.parentId);
			return Boolean(root && !root.parentId && !this.deleted.get(root.id));
		};
		return [...this.records.values()]
			.filter(visible)
			.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0))
			.map((comment) => ({
				...comment,
				...(this.resolved.has(comment.id) ? { resolved: this.resolved.get(comment.id)! } : {}),
			}));
	}

	onChange(listener: () => void): () => void {
		const types = [this.records, this.resolved, this.deleted];
		for (const type of types) type.observe(listener);
		return () => {
			for (const type of types) type.unobserve(listener);
		};
	}

	private change(action: () => void): boolean {
		if (!this.writable() || !this.session.canWrite()) return false;
		this.stopCapturing();
		try {
			this.session.doc.transact(action, ySyncPluginKey);
		} finally {
			this.stopCapturing();
		}
		return true;
	}

	add(
		view: EditorView,
		author: string,
		text: string,
		idGenerator: (kind: string) => string,
	): Comment | null {
		if (
			!view.editable ||
			view.state.selection.empty ||
			ySyncPluginKey.getState(view.state)?.type !== this.fragment
		)
			return null;
		const id = idGenerator('comment');
		if (!id || this.records.has(id)) return null;
		let added: Comment | null = null;
		this.change(() => {
			added = addComment(view, author, text, () => id);
			if (added) this.records.set(id, added);
		});
		return added;
	}

	reply(
		parentId: string,
		author: string,
		text: string,
		idGenerator: (kind: string) => string,
	): boolean {
		const root = this.records.get(parentId);
		if (!root || root.parentId || this.deleted.get(parentId)) return false;
		const id = idGenerator('comment');
		if (!id || this.records.has(id)) return false;
		return this.change(() => this.records.set(id, { id, parentId, author, text }));
	}

	resolve(id: string, resolved: boolean): boolean {
		const root = this.records.get(id);
		if (!root || root.parentId || this.deleted.get(id)) return false;
		return this.change(() => this.resolved.set(id, resolved));
	}

	delete(view: EditorView, id: string): boolean {
		if (
			!view.editable ||
			ySyncPluginKey.getState(view.state)?.type !== this.fragment ||
			!this.records.has(id) ||
			this.deleted.get(id)
		)
			return false;
		return this.change(() => {
			removeSharedAnchor(this.fragment, id);
			this.deleted.set(id, true);
		});
	}
}

/** Format only this anchor's Yjs attribute. Replacing a PM mark set also
 * reasserts the remaining anchors, which can revive a concurrent deletion. */
function removeSharedAnchor(fragment: Y.XmlFragment, id: string): void {
	for (const child of fragment.toArray()) {
		if (child instanceof Y.XmlElement) removeSharedAnchor(child, id);
		else if (child instanceof Y.XmlText) {
			let offset = 0;
			for (const part of child.toDelta()) {
				const length = typeof part.insert === 'string' ? part.insert.length : 1;
				for (const [key, attrs] of Object.entries(part.attributes ?? {})) {
					if (
						key.startsWith('comment--') &&
						Array.isArray((attrs as { ids?: unknown }).ids) &&
						(attrs as { ids: unknown[] }).ids.includes(id)
					)
						child.format(offset, length, { [key]: null });
				}
				offset += length;
			}
		}
	}
}
