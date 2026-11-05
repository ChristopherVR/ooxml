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
		// IndexedDB structured-clones the record when it is put, so no defensive copy is needed here.
		baseline: source.baseline,
		pending,
	};
}

/** A connected document without its bytes: what a lookup by cloud identity needs. */
export interface SuiteCloudRecord {
	id: string;
	name: string;
	kind: SuiteFileKind;
	revision: number;
	cloud: Omit<SuiteCloudSource, 'baseline'>;
}
const CLOUD_STORE = 'cloud';
const cloudRecord = (doc: SuiteDocument): SuiteCloudRecord | undefined => {
	if (!doc.cloud) return undefined;
	const { baseline: _baseline, ...cloud } = doc.cloud;
	return { id: doc.id, name: doc.name, kind: doc.kind, revision: doc.revision, cloud };
};

/** Browser persistence with transactional revision checks, shared by tabs and Teams attachments. */
export class SuiteDocumentStore {
	private readonly db: Promise<IDBDatabase>;
	constructor(name = 'ooxml-suite-documents') {
		this.db = new Promise((resolve, reject) => {
			const request = indexedDB.open(name, 2);
			request.onupgradeneeded = (event) => {
				const db = request.result;
				if (!db.objectStoreNames.contains('documents'))
					db.createObjectStore('documents', { keyPath: 'id' });
				// Version 2 adds a bytes-free index of connected documents; backfill it from version 1.
				const meta = db.createObjectStore(CLOUD_STORE, { keyPath: 'id' });
				if (event.oldVersion >= 1) {
					const cursor = request.transaction!.objectStore('documents').openCursor();
					cursor.onsuccess = () => {
						const entry = cursor.result;
						if (!entry) return;
						const record = cloudRecord(entry.value as SuiteDocument);
						if (record) meta.put(record);
						entry.continue();
					};
				}
			};
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
	/** Connected documents without reading any document bytes (see `SuiteCloudRecord`). */
	async listCloud(): Promise<SuiteCloudRecord[]> {
		const db = await this.db;
		return new Promise((resolve, reject) => {
			const request = db.transaction(CLOUD_STORE).objectStore(CLOUD_STORE).getAll();
			request.onsuccess = () => resolve(request.result as SuiteCloudRecord[]);
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
			const tx = db.transaction(['documents', CLOUD_STORE], 'readwrite');
			const records = tx.objectStore('documents');
			const meta = tx.objectStore(CLOUD_STORE);
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
				for (const doc of removed) {
					records.delete(doc.id);
					meta.delete(doc.id);
				}
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
			const tx = db.transaction(['documents', CLOUD_STORE], 'readwrite');
			const records = tx.objectStore('documents');
			const meta = tx.objectStore(CLOUD_STORE);
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
					const record = cloudRecord(document);
					if (record) meta.put(record);
					else meta.delete(document.id);
				};
			}
		});
	}
}
