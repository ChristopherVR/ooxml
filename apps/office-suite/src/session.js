import { mountEditor, previewRefusal } from './editors.js';
import { cloudReadOnly } from './cloud-files.js';
import { isLocked, lockInput, takeComposed, takeQueued } from './session-lock.js';
import { deepActive, focusPlace, holds, returnFocus } from './focus-place.js';
import { displayLanguage, onSettings } from './settings.js';

/**
 * One live editor per document id. Save operations are serialized per session. A session whose
 * editor shows the assistant's suggestions (`preview`, see {@link SuiteSessions.setPreview}) is
 * never saved, exported, printed or copied out until the review applies, discards or keeps them.
 */
export class SuiteSessions {
	constructor(store, host, update, command) {
		Object.assign(this, { store, host, update, command });
		this.sessions = new Map();
		// The editors show the display language, whatever the workspace around them is written in.
		onSettings((_settings, keys) => {
			if (!keys.includes('locale')) return;
			for (const { node } of this.sessions.values()) Object.assign(node, displayLanguage());
		});
	}
	async open(doc) {
		if (this.sessions.has(doc.id)) return this.sessions.get(doc.id).ready;
		const node = document.createElement('section');
		node.className = 'document-host';
		node.setAttribute('aria-label', doc.name);
		Object.assign(node, displayLanguage());
		this.host.append(node);
		const session = { doc, node, version: 0, saved: 0, loading: true, pending: Promise.resolve() };
		this.sessions.set(doc.id, session);
		session.ready = mountEditor(
			node,
			doc,
			() => {
				if (!session.loading) {
					session.version++;
					this.update();
				}
			},
			(command, detail) => this.command(command, doc.id, detail),
			// The slide or page in view, or the selection's formatting, changed (Simple view's bar).
			() => {
				if (!session.loading) this.moved?.();
			},
		)
			.then(async (editor) => {
				session.editor = editor;
				session.cloudReadOnly = await this.readOnly(doc);
				editor.setCloudReadOnly?.(session.cloudReadOnly);
				session.loading = false;
				return session;
			})
			.catch((error) => {
				node.remove();
				this.sessions.delete(doc.id);
				throw error;
			});
		return session.ready;
	}
	async save(id, { cloud = true } = {}) {
		const s = await this.sessions.get(id)?.ready;
		if (!s) return this.store.get(id);
		// The assistant's suggestions on screen are a preview: only their review panel saves them.
		if (s.preview) throw previewRefusal(s.doc.name);
		const run = async () => {
			await s.editor.flush?.();
			if (s.version === s.saved) {
				const latest = await this.store.get(id);
				if (latest.revision !== s.doc.revision) await this.reload(latest);
				return cloud ? this.saveRemote(latest) : latest;
			}
			const version = s.version;
			if (await this.readOnly(s.doc))
				throw new Error(
					'This connected document is read-only. Download a copy to edit separately.',
				);
			const bytes = await s.editor.save();
			if (s.version !== version)
				throw new Error('The document changed during save. Please save again.');
			const changes = [];
			let next = {
				...s.doc,
				bytes,
				revision: s.doc.revision + 1,
				modified: Date.now(),
				...(s.doc.cloud ? { cloud: { ...s.doc.cloud, pending: true } } : {}),
			};
			const child = next;
			while (next.parent) {
				const p = next.parent;
				const parent = await this.store.get(p.id);
				if (this.isEdited(p.id))
					throw new Error('Save the parent document before saving its embedded document.');
				if (parent.revision !== p.revision)
					throw new Error('The parent changed. Reopen the embedded document before saving.');
				const { replaceEmbeddedPackage } = await import('ooxml-core/opc');
				const bytes = await replaceEmbeddedPackage(parent.bytes, p.path, next.bytes);
				next.parent = { ...p, revision: parent.revision + 1 };
				changes.push({ document: next, expected: next.revision - 1 });
				next = {
					...parent,
					bytes,
					revision: parent.revision + 1,
					modified: Date.now(),
					...(parent.cloud ? { cloud: { ...parent.cloud, pending: true } } : {}),
				};
			}
			changes.push({ document: next, expected: next.revision - 1 });
			if (
				s.version !== version ||
				changes.some((c) => c.document.id !== id && this.isEdited(c.document.id))
			)
				throw new Error('A document changed while saving. Please try again.');
			const unlock = this.lock(changes.map((c) => c.document.id));
			try {
				try {
					await this.store.commit(changes);
				} catch (error) {
					if (/changed in another session/u.test(error.message))
						throw new Error(
							'This file changed in another tab. Your edits are still open. Download a copy before reopening it.',
							{ cause: error },
						);
					throw error;
				}
				s.doc = child;
				s.saved = version;
				s.editor.clean?.();
				for (const change of changes)
					if (change.document.id !== id) await this.reload(change.document);
			} finally {
				unlock();
			}
			this.update();
			if (cloud) await this.saveRemote(next);
			return this.store.get(child.id);
		};
		const result = s.pending.then(run);
		s.pending = result.catch(() => {});
		return result;
	}
	/**
	 * Marks `session` as showing the assistant's suggestions (ai/review.js) or not. While it does,
	 * the suite refuses to save or download it and its editor offers no command that takes the
	 * document out of the suite (Save As, Export, Print, Share, Copy).
	 */
	setPreview(session, on) {
		session.preview = on;
		session.editor.hold?.(on);
		session.node.toggleAttribute('data-preview', on);
	}
	/**
	 * The document as a request to the assistant reads it: the user's own edits are saved first,
	 * while a document showing the assistant's suggestions is read as last saved and its
	 * suggestions stay unsaved.
	 */
	async forContext(id) {
		const s = await this.sessions.get(id)?.ready;
		return s?.preview ? this.store.get(id) : this.save(id, { cloud: false });
	}
	isDirty(id) {
		return this.isEdited(id) || Boolean(this.sessions.get(id)?.doc.cloud?.pending);
	}
	/** Export works even when a stale local revision or cloud conflict prevents saving. */
	async export(id) {
		const session = await this.sessions.get(id)?.ready;
		if (!session) return this.store.get(id);
		if (session.preview) throw previewRefusal(session.doc.name);
		const unlock = this.lock([id]);
		try {
			await session.editor.flush?.();
			return { ...session.doc, bytes: await session.editor.save() };
		} finally {
			unlock();
		}
	}
	isEdited(id) {
		const s = this.sessions.get(id);
		return Boolean(s && s.version !== s.saved);
	}
	async readOnly(doc) {
		const seen = new Set();
		while (doc) {
			if (cloudReadOnly(doc)) return true;
			if (!doc.parent || seen.has(doc.id)) return false;
			seen.add(doc.id);
			doc = await this.store.get(doc.parent.id);
		}
		return false;
	}
	async saveRemote(doc) {
		if (!doc.cloud?.pending) return doc;
		if (!this.remote)
			throw new Error('Cloud storage is not configured. Changes are kept on this device.');
		const unlock = this.lock([doc.id]);
		try {
			const next = await this.remote.save(doc);
			const session = this.sessions.get(doc.id);
			if (session) session.doc = next;
			this.update();
			return next;
		} finally {
			unlock();
		}
	}
	/**
	 * Refuse input to these documents until the returned function runs; focus stays put. A save
	 * asked for meanwhile runs once the last lock on its document is released.
	 */
	lock(ids) {
		const locked = ids.map((id) => this.sessions.get(id)).filter(Boolean);
		const unlocks = locked.map((s) => lockInput(s.node));
		return () =>
			locked.forEach((s, index) => {
				if (!unlocks[index]()) return;
				this.settleComposed(s);
				// The suite's command callback reports a failure itself.
				for (const command of takeQueued(s.node))
					void this.command(command, s.doc.id).catch(() => {});
			});
	}
	/**
	 * Accounts for what an input method wrote while the document was locked. Text a reload threw
	 * away is typed again where the caret is; text still in the editor is an unsaved edit, so the
	 * tab shows it and the next save keeps it.
	 */
	settleComposed(s) {
		const lost = s.lostComposition;
		s.lostComposition = '';
		if (lost) {
			if (!holds(s.node, deepActive())) s.editor.focus?.();
			if (holds(s.node, deepActive())) document.execCommand('insertText', false, lost);
		}
		if (takeComposed(s.node)) {
			s.version++;
			this.update();
		}
	}
	async reload(doc) {
		const s = this.sessions.get(doc.id);
		if (!s) return;
		// Preview documents need not have been committed; derive their cloud state from the bytes.
		if (doc.cloud)
			doc = {
				...doc,
				cloud: {
					...doc.cloud,
					pending:
						doc.bytes.length !== doc.cloud.baseline.length ||
						doc.bytes.some((value, i) => value !== doc.cloud.baseline[i]),
				},
			};
		// The user stays on the same cells, caret, slide or page, and keeps keyboard focus.
		const place = focusPlace(s.node);
		const back = s.editor.keep?.();
		s.loading = true;
		try {
			await s.editor.load(doc.bytes);
			back?.(Boolean(place));
			if (place && !returnFocus(s.node, place)) s.editor.focus?.();
			// The new content lacks what an input method wrote during the lock: type it again.
			s.lostComposition = (s.lostComposition ?? '') + takeComposed(s.node);
			s.doc = doc;
			s.version = 0;
			s.saved = 0;
			s.cloudReadOnly = await this.readOnly(doc);
			s.editor.setCloudReadOnly?.(s.cloudReadOnly);
		} finally {
			s.loading = false;
		}
		if (!isLocked(s.node)) this.settleComposed(s);
	}
	/** Moves a tab to `index` among the open documents; tab order is the map's order. */
	move(id, index) {
		const s = this.sessions.get(id);
		if (!s) return;
		const order = [...this.sessions.keys()].filter((key) => key !== id);
		order.splice(Math.max(0, Math.min(index, order.length)), 0, id);
		this.sessions = new Map(order.map((key) => [key, key === id ? s : this.sessions.get(key)]));
		this.update();
	}
	close(id) {
		const s = this.sessions.get(id);
		if (!s) return;
		s.editor?.dispose();
		s.node.remove();
		this.sessions.delete(id);
	}
}
