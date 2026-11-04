/** One cloud connection for every Office format. Authentication stays with the embedding host. */
export function cloudIdentity(file) {
	return Object.fromEntries(
		['driveId', 'itemId', 'parentId', 'name', 'eTag', 'kind', 'access']
			.filter((key) => file[key] !== undefined)
			.map((key) => [key, file[key]]),
	);
}

export const cloudReadOnly = (doc) => Boolean(doc.cloud && doc.cloud.file.access !== 'write');
export const cloudSaveLabel = (doc, edited = false) =>
	edited
		? 'Unsaved changes'
		: doc?.cloud?.pending
			? 'Saved on this device; OneDrive save pending'
			: doc?.cloud
				? 'Saved to OneDrive'
				: 'Saved to this device';

export class SuiteOneDrive {
	constructor({ store, client, createSession, merge, conflict, adopted, changed, dirty }) {
		Object.assign(this, { store, client, createSession, merge, conflict, adopted, changed, dirty });
	}
	configured() {
		return Boolean(this.client);
	}
	connection(doc) {
		if (!this.client)
			throw new Error('OneDrive sign-in is not configured. Your local copy is kept.');
		return this.createSession(
			this.client,
			{ file: doc.cloud.file, bytes: doc.cloud.baseline },
			this.merge ? { merge: this.merge } : undefined,
		);
	}
	async open(link) {
		if (!this.client) throw new Error('OneDrive sign-in is not configured by this host.');
		const opened = await this.client.open(link);
		const existing = (await this.store.list()).find(
			(doc) =>
				doc.cloud?.file.driveId === opened.file.driveId &&
				doc.cloud?.file.itemId === opened.file.itemId,
		);
		// A repeated link must never replace an unsynchronized local document.
		if (existing && (existing.cloud.pending || this.dirty?.(existing.id))) return existing;
		const cloud = {
			file: cloudIdentity(opened.file),
			baseline: opened.bytes.slice(),
			pending: false,
		};
		if (!existing) return this.store.create(opened.file.name, opened.bytes, { cloud });
		const next = {
			...existing,
			name: opened.file.name,
			bytes: opened.bytes,
			cloud,
			revision: existing.revision + 1,
			modified: Date.now(),
		};
		await this.store.commit([{ document: next, expected: existing.revision }]);
		await this.adopted?.(next);
		return next;
	}
	async persist(doc, opened, reload = false) {
		const latest = await this.store.get(doc.id);
		if (latest.revision !== doc.revision)
			throw new Error('The local document changed while saving. Your local copy is kept.');
		const next = {
			...doc,
			name: opened.file.name,
			bytes: opened.bytes,
			cloud: { file: cloudIdentity(opened.file), baseline: opened.bytes.slice(), pending: false },
			// Cloud acknowledgement changes metadata, not package content or embedding revisions.
			revision: reload ? doc.revision + 1 : doc.revision,
			modified: Date.now(),
		};
		await this.store.commit([{ document: next, expected: doc.revision }]);
		if (reload) await this.adopted?.(next);
		this.changed?.(next);
		return next;
	}
	async save(doc) {
		if (!doc.cloud?.pending) return doc;
		const connection = this.connection(doc);
		try {
			const file = await connection.save(doc.bytes);
			const bytes = connection.getSnapshot?.().baseBytes ?? doc.bytes;
			const combined =
				bytes.length !== doc.bytes.length || bytes.some((value, i) => value !== doc.bytes[i]);
			return await this.persist(doc, { file, bytes }, combined);
		} catch (error) {
			if (!error?.conflict) {
				const messages = {
					authentication:
						'Sign in to Microsoft again. Your edits are saved on this device and remain pending.',
					'access-denied':
						'OneDrive no longer allows this save. Your edits are saved on this device; download a copy or ask the owner for access.',
					locked:
						'This OneDrive file is locked by another editor. Your edits are saved on this device. Try saving again after it is unlocked.',
					'too-large':
						'This file exceeds the OneDrive connection size limit. Your edited copy is kept on this device.',
				};
				// oxlint-disable-next-line eslint/preserve-caught-error -- Provider failures may expose credentials.
				throw new Error(
					messages[error?.code] ??
						'OneDrive save failed. Your changes are saved on this device; check the connection and try again.',
				);
			}
			const action = await this.conflict?.({ doc, conflict: error.conflict });
			if (action === 'copy') return this.persist(doc, await connection.saveCopy(doc.bytes));
			if (action === 'latest') {
				if (error.conflict.remote.file.kind !== doc.kind)
					// oxlint-disable-next-line eslint/preserve-caught-error -- Conflict causes include private document bytes.
					throw new Error(
						'The OneDrive file changed format. Save a separate copy of your edits before reopening it.',
					);
				// Keep an independent library copy before replacing the working document.
				await this.store.create(
					doc.name.replace(/(\.[^.]+)$/u, ' (local edits)$1'),
					doc.bytes.slice(),
				);
				return this.persist(doc, await connection.reload(), true);
			}
			// oxlint-disable-next-line eslint/preserve-caught-error -- Conflict details remain in the recovery flow.
			throw new Error(
				'OneDrive has a newer version. Your edits are kept on this device and remain pending.',
			);
		}
	}
	async share(doc, options) {
		if (!doc.cloud) throw new Error('This document is not connected to OneDrive.');
		if (doc.cloud.pending) throw new Error('Save or resolve the OneDrive changes before sharing.');
		if (!this.client) throw new Error('OneDrive sign-in is not configured.');
		if (options.type === 'existing') {
			const opened = await this.client.read(doc.cloud.file);
			if (!opened.file.webUrl) throw new Error('The existing OneDrive link is unavailable.');
			return opened.file.webUrl;
		}
		return this.client.share(doc.cloud.file, options);
	}
}
