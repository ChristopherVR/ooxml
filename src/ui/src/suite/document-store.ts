export type SuiteFileKind = 'docx' | 'xlsx' | 'pptx' | 'vsdx';
/** Host-serialized safe identity/revision fields for a connected provider. */
export interface SuiteCloudFile {
	[key: string]: unknown;
	name: string;
	kind: SuiteFileKind;
	access: 'read' | 'write' | 'unknown';
	driveId?: string;
	itemId?: string;
	parentId?: string;
	eTag?: string;
}
/** Stable identifiers only. Capability links and authentication tokens must never be persisted. */
export interface SuiteCloudSource {
	/** Missing on legacy records, which belong to OneDrive. Use a stable provider/account namespace. */
	providerId?: string;
	providerLabel?: string;
	file: SuiteCloudFile;
	baseline: Uint8Array;
	pending: boolean;
}
export interface SuiteDocument {
	id: string;
	name: string;
	kind: SuiteFileKind;
	bytes: Uint8Array;
	revision: number;
	modified: number;
	parent?: { id: string; path: string; revision: number };
	sourceUrl?: string;
	cloud?: SuiteCloudSource;
}

export function suiteFileKind(name: string): SuiteFileKind {
	const ext = name.split('.').pop()?.toLowerCase();
	if (ext === 'docx' || ext === 'xlsx' || ext === 'pptx' || ext === 'vsdx') return ext;
	throw new Error('Choose a .docx, .xlsx, .pptx or .vsdx file');
}

/** Derive synchronization state for every writer, including assistant edits and version restores. */
function cloudSource(source: SuiteCloudSource, bytes: Uint8Array): SuiteCloudSource {
	const { driveId, itemId, parentId, name, eTag, kind, access } = source.file;
	const pending =
		bytes.length !== source.baseline.length ||
		bytes.some((value, i) => value !== source.baseline[i]);
	if (pending && access !== 'write') throw new Error('This connected document is read-only.');
	const file: SuiteCloudFile =
		source.providerId && source.providerId !== 'onedrive'
			? structuredClone(source.file)
			: {
					name,
					kind,
					access,
					...(driveId ? { driveId } : {}),
					...(itemId ? { itemId } : {}),
					...(eTag ? { eTag } : {}),
					...(parentId ? { parentId } : {}),
				};
	// A custom provider must serialize safe metadata; also strip known credential/link fields here.
	for (const key of ['webUrl', 'accessToken', 'refreshToken', 'downloadUrl']) delete file[key];
	return {
		...(source.providerId ? { providerId: source.providerId } : {}),
		...(source.providerLabel ? { providerLabel: source.providerLabel } : {}),
		file,
		baseline: source.baseline.slice(),
		pending,
	};
}

/** Browser persistence with transactional revision checks, shared by tabs and Teams attachments. */
export class SuiteDocumentStore {
	private readonly db: Promise<IDBDatabase>;
	constructor(name = 'ooxml-suite-documents') {
		this.db = new Promise((resolve, reject) => {
			const request = indexedDB.open(name, 1);
			request.onupgradeneeded = () =>
				request.result.createObjectStore('documents', { keyPath: 'id' });
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
	}
	async list(): Promise<SuiteDocument[]> {
		const db = await this.db;
		return new Promise((resolve, reject) => {
			const request = db.transaction('documents').objectStore('documents').getAll();
			request.onsuccess = () => resolve(request.result as SuiteDocument[]);
			request.onerror = () => reject(request.error);
		});
	}
	async get(id: string): Promise<SuiteDocument> {
		const db = await this.db;
		return new Promise((resolve, reject) => {
			const request = db.transaction('documents').objectStore('documents').get(id);
			request.onsuccess = () =>
				request.result
					? resolve(request.result as SuiteDocument)
					: reject(new Error('Document not found'));
			request.onerror = () => reject(request.error);
		});
	}
	async create(
		name: string,
		bytes: Uint8Array,
		extra: Pick<SuiteDocument, 'parent' | 'sourceUrl' | 'cloud'> = {},
	): Promise<SuiteDocument> {
		const doc: SuiteDocument = {
			id: crypto.randomUUID(),
			name,
			kind: suiteFileKind(name),
			bytes,
			revision: 1,
			modified: Date.now(),
			...extra,
		};
		await this.commit([{ document: doc, expected: 0 }]);
		return doc;
	}
	/** Remove one library copy and cached embedded children atomically. Returns snapshots for Undo. */
	async remove(id: string, expected: number): Promise<SuiteDocument[]> {
		const db = await this.db;
		return new Promise((resolve, reject) => {
			const tx = db.transaction('documents', 'readwrite');
			const records = tx.objectStore('documents');
			let removed: SuiteDocument[] = [],
				failure: Error | undefined;
			tx.oncomplete = () => resolve(removed);
			tx.onabort = () => reject(failure ?? tx.error ?? new Error('Removal failed'));
			const request = records.getAll();
			request.onsuccess = () => {
				const all = request.result as SuiteDocument[];
				if (all.find((d) => d.id === id)?.revision !== expected) {
					failure = new Error('This file changed in another session. Refresh before removing it.');
					tx.abort();
					return;
				}
				const ids = new Set([id]);
				let count = 0;
				while (count !== ids.size) {
					count = ids.size;
					for (const doc of all) if (doc.parent && ids.has(doc.parent.id)) ids.add(doc.id);
				}
				removed = all.filter((doc) => ids.has(doc.id));
				for (const doc of removed) records.delete(doc.id);
			};
		});
	}

	/** All replacements succeed together, including child and parent when saving an embedding. */
	async commit(changes: { document: SuiteDocument; expected: number }[]): Promise<void> {
		changes = changes.map(({ document, expected }) => ({
			document: document.cloud
				? { ...document, cloud: cloudSource(document.cloud, document.bytes) }
				: document,
			expected,
		}));
		const db = await this.db;
		return new Promise((resolve, reject) => {
			const tx = db.transaction('documents', 'readwrite');
			const records = tx.objectStore('documents');
			let failure: Error | undefined;
			tx.oncomplete = () => resolve();
			tx.onabort = () => reject(failure ?? tx.error ?? new Error('Save failed'));
			for (const { document, expected } of changes) {
				const request = records.get(document.id);
				request.onsuccess = () => {
					if (((request.result as SuiteDocument | undefined)?.revision ?? 0) !== expected) {
						failure = new Error(
							`${document.name} changed in another session. Reopen it before saving.`,
						);
						tx.abort();
						return;
					}
					records.put(document);
				};
			}
		});
	}
}
