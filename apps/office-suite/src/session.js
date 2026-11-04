import { replaceEmbeddedPackage } from 'ooxml-core/opc';
import { mountEditor } from './editors.js';
import { cloudReadOnly } from './cloud-files.js';

/** One live editor per document id. Save operations are serialized per session. */
export class SuiteSessions {
	constructor(store, host, update, command) {
		Object.assign(this, { store, host, update, command });
		this.sessions = new Map();
	}
	async open(doc) {
		if (this.sessions.has(doc.id)) return this.sessions.get(doc.id).ready;
		const node = document.createElement('section');
		node.className = 'document-host';
		node.setAttribute('aria-label', doc.name);
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
			(command) => this.command(command, doc.id),
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
		const run = async () => {
			s.editor.flush?.();
			if (s.version === s.saved) {
				const latest = await this.store.get(id);
				if (latest.revision !== s.doc.revision) await this.reload(latest);
				return cloud ? this.saveRemote(latest) : latest;
			}
			const version = s.version;
			if (await this.readOnly(s.doc))
				throw new Error('This OneDrive document is read-only. Download a copy to edit separately.');
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
	isDirty(id) {
		return this.isEdited(id) || Boolean(this.sessions.get(id)?.doc.cloud?.pending);
	}
	/** Export works even when a stale local revision or cloud conflict prevents saving. */
	async export(id) {
		const session = await this.sessions.get(id)?.ready;
		if (!session) return this.store.get(id);
		const unlock = this.lock([id]);
		try {
			session.editor.flush?.();
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
			throw new Error('OneDrive is not configured. Changes are kept on this device.');
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
	lock(ids) {
		const nodes = ids.map((id) => this.sessions.get(id)?.node).filter(Boolean);
		const states = nodes.map((node) => node.inert);
		nodes.forEach((node) => {
			node.inert = true;
		});
		return () =>
			nodes.forEach((node, index) => {
				node.inert = states[index];
			});
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
		s.loading = true;
		try {
			await s.editor.load(doc.bytes);
			s.doc = doc;
			s.version = 0;
			s.saved = 0;
			s.cloudReadOnly = await this.readOnly(doc);
			s.editor.setCloudReadOnly?.(s.cloudReadOnly);
		} finally {
			s.loading = false;
		}
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
