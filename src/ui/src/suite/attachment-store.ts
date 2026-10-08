export interface SuiteAttachment {
	id: string;
	name: string;
	mime: string;
	bytes: Uint8Array;
}

/** General file storage for chat attachments, independent of the Office document model. */
export class SuiteAttachmentStore {
	private readonly db: Promise<IDBDatabase>;
	constructor(name = 'ooxml-suite-attachments') {
		this.db = new Promise((resolve, reject) => {
			const request = indexedDB.open(name, 1);
			request.onupgradeneeded = () => request.result.createObjectStore('files', { keyPath: 'id' });
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
	}
	async create(name: string, mime: string, bytes: Uint8Array): Promise<SuiteAttachment> {
		const file = { id: crypto.randomUUID(), name, mime, bytes };
		const db = await this.db;
		await new Promise<void>((resolve, reject) => {
			const tx = db.transaction('files', 'readwrite');
			tx.objectStore('files').put(file);
			tx.oncomplete = () => resolve();
			tx.onabort = () => reject(tx.error ?? new Error('The attachment could not be stored.'));
		});
		return file;
	}
	async get(id: string): Promise<SuiteAttachment> {
		const db = await this.db;
		return new Promise((resolve, reject) => {
			const request = db.transaction('files').objectStore('files').get(id);
			request.onsuccess = () =>
				request.result
					? resolve(request.result as SuiteAttachment)
					: reject(new Error('This attachment is not available in the current local profile.'));
			request.onerror = () => reject(request.error);
		});
	}
}
