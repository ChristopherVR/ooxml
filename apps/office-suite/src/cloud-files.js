/** One cloud connection for every Office format. Authentication stays with the embedding host. */
export function cloudIdentity(file) {
	return Object.fromEntries(
		['driveId', 'itemId', 'parentId', 'name', 'eTag', 'kind', 'access']
			.filter((key) => file[key] !== undefined)
			.map((key) => [key, file[key]]),
	);
}

export const cloudReadOnly = (doc) => Boolean(doc.cloud && doc.cloud.file.access !== 'write');
export const cloudProviderLabel = (doc) => doc?.cloud?.providerLabel ?? 'OneDrive';
export const cloudSaveLabel = (doc, edited = false) =>
	edited
		? 'Unsaved changes'
		: doc?.cloud?.pending
			? `Saved on this device; ${cloudProviderLabel(doc)} save pending`
			: doc?.cloud
				? `Saved to ${cloudProviderLabel(doc)}`
				: 'Saved to this device';

export class SuiteCloudStorage {
	constructor({
		store,
		client,
		createSession,
		merge,
		conflict,
		adopted,
		changed,
		dirty,
		stale,
		notice,
		progress,
		id = 'onedrive',
		label = 'OneDrive',
		serializeFile = cloudIdentity,
		keyOf = (file) => JSON.stringify([file.driveId, file.itemId]),
		sessionOptions = {},
	}) {
		Object.assign(this, {
			store,
			client,
			createSession,
			merge,
			conflict,
			adopted,
			changed,
			dirty,
			stale,
			notice,
			progress,
			id,
			label,
			serializeFile,
			keyOf,
			sessionOptions,
		});
	}
	configured() {
		return Boolean(this.client);
	}
	connection(doc) {
		this.checkProvider(doc);
		if (!this.client)
			throw this.error('OneDrive sign-in is not configured. Your local copy is kept.');
		return this.createSession(
			this.client,
			{ file: doc.cloud.file, bytes: doc.cloud.baseline },
			{ ...this.sessionOptions, merge: this.merge },
		);
	}
	/** The library document already connected to this cloud file, without reading every document. */
	async find(file) {
		const key = this.keyOf(file);
		// listCloud() is a bytes-free index; stores without it fall back to loading every document.
		const records = this.store.listCloud ? await this.store.listCloud() : await this.store.list();
		const match = records.find(
			(doc) =>
				doc.cloud &&
				(doc.cloud.providerId ?? 'onedrive') === this.id &&
				this.keyOf(doc.cloud.file) === key,
		);
		return match ? this.store.get(match.id) : undefined;
	}
	async open(link) {
		if (!this.client) throw this.error('OneDrive sign-in is not configured by this host.');
		return this.adopt(await this.client.open(link));
	}
	/** Adds an opened cloud file ({ file, bytes }) to the library, or reuses its existing document. */
	async adopt(opened) {
		const existing = await this.find(opened.file);
		// Edits on this device that OneDrive does not have yet are never replaced silently.
		if (existing && (existing.cloud.pending || this.dirty?.(existing.id))) {
			const unchanged = existing.cloud.file?.eTag === opened.file.eTag;
			const choice = unchanged
				? 'keep'
				: ((await this.stale?.({ doc: existing, remote: opened })) ?? 'keep');
			if (choice === 'keep') {
				this.notice?.(
					unchanged
						? `Reopened your edits from this device. Save to update ${this.label}.`
						: `Kept your edits from this device. ${this.label} has a newer version; Save combines them where possible.`,
				);
				return existing;
			}
			// Open the newer cloud version, keeping the local edits as a separate library file.
			await this.store.create(
				existing.name.replace(/(\.[^.]+)$/u, ' (local edits)$1'),
				existing.bytes,
			);
		}
		const cloud = {
			providerId: this.id,
			providerLabel: this.label,
			file: this.serializeFile(opened.file),
			// The store commits a structured clone, so baseline and bytes may share one download.
			baseline: opened.bytes,
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
			throw this.error('The local document changed while saving. Your local copy is kept.');
		const next = {
			...doc,
			name: opened.file.name,
			bytes: opened.bytes,
			cloud: {
				providerId: this.id,
				providerLabel: this.label,
				file: this.serializeFile(opened.file),
				baseline: opened.bytes,
				pending: false,
			},
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
		try {
			return await this.upload(doc);
		} finally {
			this.progress?.(doc, null);
		}
	}
	async upload(doc) {
		const connection = this.connection(doc);
		// Each step of the upload (and of a merge) for the host's status; null when it ends.
		const onSaveProgress = (step) => this.progress?.(doc, step);
		try {
			const file = await connection.save(doc.bytes, { onSaveProgress });
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
				throw this.error(
					messages[error?.code] ??
						'OneDrive save failed. Your changes are saved on this device; check the connection and try again.',
				);
			}
			const action = await this.conflict?.({ doc, conflict: error.conflict });
			if (action === 'copy') return this.persist(doc, await connection.saveCopy(doc.bytes));
			if (action === 'latest') {
				if (error.conflict.remote.file.kind !== doc.kind)
					// oxlint-disable-next-line eslint/preserve-caught-error -- Conflict causes include private document bytes.
					throw this.error(
						'The OneDrive file changed format. Save a separate copy of your edits before reopening it.',
					);
				// Keep an independent library copy before replacing the working document.
				await this.store.create(doc.name.replace(/(\.[^.]+)$/u, ' (local edits)$1'), doc.bytes);
				return this.persist(doc, await connection.reload(), true);
			}
			// oxlint-disable-next-line eslint/preserve-caught-error -- Conflict details remain in the recovery flow.
			throw this.error(
				'OneDrive has a newer version. Your edits are kept on this device and remain pending.',
			);
		}
	}
	async share(doc, options) {
		this.checkProvider(doc);
		if (!doc.cloud) throw this.error('This document is not connected to OneDrive.');
		if (!this.client) throw this.error('OneDrive sign-in is not configured.');
		if (options.type === 'existing') {
			// Only the link is needed: ask for the item's metadata, not its whole content.
			const current = this.client.metadata
				? await this.client.metadata(doc.cloud.file)
				: (await this.client.read(doc.cloud.file)).file;
			if (!current.webUrl) throw this.error('The existing OneDrive link is unavailable.');
			return current.webUrl;
		}
		if (!this.client.share)
			throw this.error(`${this.label} does not support creating share links.`);
		return this.client.share(doc.cloud.file, options);
	}
	error(text) {
		return new Error(
			text
				.replaceAll('OneDrive', () => this.label)
				.replaceAll('Microsoft', () => (this.label === 'OneDrive' ? 'Microsoft' : this.label)),
		);
	}
	checkProvider(doc) {
		if (doc.cloud && (doc.cloud.providerId ?? 'onedrive') !== this.id)
			throw this.error(
				'Connect the storage provider that owns this document. Your local copy is kept.',
			);
	}
}

/** Compatibility name for existing suite hosts. */
export { SuiteCloudStorage as SuiteOneDrive };
