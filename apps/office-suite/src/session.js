import { replaceEmbeddedPackage } from 'ooxml-core/opc';
import { mountEditor } from './editors.js';

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
			.then((editor) => {
				session.editor = editor;
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
	async save(id) {
		const s = await this.sessions.get(id)?.ready;
		if (!s) return this.store.get(id);
		const run = async () => {
			s.editor.flush?.();
			if (s.version === s.saved) {
				const latest = await this.store.get(id);
				if (latest.revision !== s.doc.revision) await this.reload(latest);
				return latest;
			}
			const version = s.version;
			const bytes = await s.editor.save();
			if (s.version !== version)
				throw new Error('The document changed during save. Please save again.');
			const changes = [];
			let next = { ...s.doc, bytes, revision: s.doc.revision + 1, modified: Date.now() };
			const child = next;
			while (next.parent) {
				const p = next.parent;
				const parent = await this.store.get(p.id);
				if (this.isDirty(p.id))
					throw new Error('Save the parent document before saving its embedded document.');
				if (parent.revision !== p.revision)
					throw new Error('The parent changed. Reopen the embedded document before saving.');
				const bytes = await replaceEmbeddedPackage(parent.bytes, p.path, next.bytes);
				next.parent = { ...p, revision: parent.revision + 1 };
				changes.push({ document: next, expected: next.revision - 1 });
				next = { ...parent, bytes, revision: parent.revision + 1, modified: Date.now() };
			}
			changes.push({ document: next, expected: next.revision - 1 });
			if (
				s.version !== version ||
				changes.some((c) => c.document.id !== id && this.isDirty(c.document.id))
			)
				throw new Error('A document changed while saving. Please try again.');
			const unlock = this.lock(changes.map((c) => c.document.id));
			try {
				await this.store.commit(changes);
				s.doc = child;
				s.saved = version;
				s.editor.clean?.();
				for (const change of changes)
					if (change.document.id !== id) await this.reload(change.document);
			} finally {
				unlock();
			}
			this.update();
			return child;
		};
		const result = s.pending.then(run);
		s.pending = result.catch(() => {});
		return result;
	}
	isDirty(id) {
		const s = this.sessions.get(id);
		return Boolean(s && s.version !== s.saved);
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
		s.loading = true;
		try {
			await s.editor.load(doc.bytes);
			s.doc = doc;
			s.version = 0;
			s.saved = 0;
		} finally {
			s.loading = false;
		}
	}
	close(id) {
		const s = this.sessions.get(id);
		if (!s) return;
		s.editor?.dispose();
		s.node.remove();
		this.sessions.delete(id);
	}
}
