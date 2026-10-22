import { recoveryStoreOf, type DocumentStorage } from 'ooxml-core/embedding';
import type { ViewerController } from './controller';

/** A drawing with unsaved changes, kept as the bytes of the .vsdx it would save as. */
export interface RecoveryEntry {
	/** The file name it came from: the key, so reopening the same file finds it. */
	name: string;
	savedAt: number;
	bytes: Uint8Array;
}

/** Where recovery copies live. Every method may fail; callers treat a failure as "nothing kept". */
export interface RecoveryStore {
	get(name: string): Promise<RecoveryEntry | undefined>;
	put(entry: RecoveryEntry): Promise<void>;
	delete(name: string): Promise<void>;
}

/** A store in memory, for tests and for hosts without browser storage. */
export function memoryRecoveryStore(): RecoveryStore & { entries: Map<string, RecoveryEntry> } {
	const entries = new Map<string, RecoveryEntry>();
	return {
		entries,
		get: async (name) => entries.get(name),
		put: async (entry) => void entries.set(entry.name, { ...entry, bytes: entry.bytes.slice() }),
		delete: async (name) => void entries.delete(name),
	};
}

const DATABASE = 'ooxml-visio-recovery';
const STORE = 'drawings';

const settle = <T>(request: IDBRequest<T>): Promise<T> =>
	new Promise((resolve, reject) => {
		request.onsuccess = () => resolve(request.result);
		request.onerror = () => reject(request.error);
	});

/** The browser's IndexedDB as a recovery store, or undefined where there is none. */
export function indexedDbRecoveryStore(
	factory: IDBFactory | undefined = globalThis.indexedDB,
): RecoveryStore | undefined {
	if (!factory) return undefined;
	let opened: Promise<IDBDatabase> | undefined;
	const database = (): Promise<IDBDatabase> =>
		(opened ??= new Promise((resolve, reject) => {
			const request = factory.open(DATABASE, 1);
			request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: 'name' });
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		}));
	const run = async <T>(mode: IDBTransactionMode, work: (store: IDBObjectStore) => IDBRequest<T>) =>
		settle(work((await database()).transaction(STORE, mode).objectStore(STORE)));
	return {
		get: async (name) =>
			(await run('readonly', (store) => store.get(name))) as RecoveryEntry | undefined,
		put: async (entry) => void (await run('readwrite', (store) => store.put(entry))),
		delete: async (name) => void (await run('readwrite', (store) => store.delete(name))),
	};
}

/**
 * A host's `DocumentStorage` recovery methods as a recovery store, or undefined when the storage
 * has none (the viewer then keeps using IndexedDB). The drawing's file name is the document id.
 */
export function storageRecoveryStore(
	storage: DocumentStorage | null | undefined,
): RecoveryStore | undefined {
	const recovery = recoveryStoreOf(storage);
	if (!recovery) return undefined;
	return {
		get: async (name) => {
			const entry = await recovery.readRecovery(name);
			return entry && { name: entry.id, savedAt: entry.savedAt, bytes: entry.bytes };
		},
		put: (entry) =>
			recovery.writeRecovery({ id: entry.name, savedAt: entry.savedAt, bytes: entry.bytes }),
		delete: (name) => recovery.clearRecovery(name),
	};
}

/** A kept copy found when a drawing opens; the host asks whether to take it back. */
export interface RecoveryOffer {
	name: string;
	savedAt: number;
	/** Open the kept copy in place of the file that was just opened. */
	restore(): Promise<void>;
	/** Throw the kept copy away. */
	discard(): Promise<void>;
}

interface Timers {
	setInterval(handler: () => void, ms: number): unknown;
	clearInterval(id: unknown): void;
}

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean =>
	a.byteLength === b.byteLength && a.every((value, index) => value === b[index]);

/**
 * File > Options > Save > AutoRecover. While it is on, a drawing with unsaved changes is copied to
 * the store every few minutes and when the page is hidden. Opening a file whose name has a kept
 * copy offers it back, and saving a copy (File > Save) removes it, since the copy then exists as a
 * file. A kept copy is offered even after the option is switched off: the data is already there.
 */
export class ViewerAutoRecover {
	#enabled = false;
	#timer: unknown;
	#lookup = 0;
	constructor(
		private readonly controller: ViewerController,
		private readonly fileName: () => string,
		/** The store, or a getter read at each use so a host can swap it (the `storage` property). */
		store: RecoveryStore | undefined | (() => RecoveryStore | undefined),
		private readonly offer: (offer: RecoveryOffer) => void,
		private readonly now: () => number = () => Date.now(),
		private readonly timers: Timers = globalThis as unknown as Timers,
	) {
		this.#store = typeof store === 'function' ? store : () => store;
	}
	readonly #store: () => RecoveryStore | undefined;
	private get store(): RecoveryStore | undefined {
		return this.#store();
	}
	/** Apply the option: copy every `minutes` minutes while `enabled`. */
	configure(enabled: boolean, minutes: number): void {
		if (this.#timer !== undefined) this.timers.clearInterval(this.#timer);
		this.#timer = undefined;
		this.#enabled = enabled;
		if (this.#enabled)
			this.#timer = this.timers.setInterval(() => void this.saveNow(), minutes * 60_000);
	}
	get enabled(): boolean {
		return this.#enabled && !!this.store;
	}
	#name(): string {
		return this.fileName() || 'Drawing1.vsdx';
	}
	/** Copy the drawing now if it has unsaved changes; true when a copy was kept. */
	async saveNow(): Promise<boolean> {
		const state = this.controller.state;
		if (
			!this.#enabled ||
			!this.store ||
			!state.edit.dirty ||
			!state.edit.sourceAvailable ||
			state.loading ||
			state.edit.busy
		)
			return false;
		try {
			await this.store.put({
				name: this.#name(),
				savedAt: this.now(),
				bytes: this.controller.exportVsdx().bytes,
			});
			return true;
		} catch {
			return false;
		}
	}
	/** A copy of the drawing was saved as a file: the kept one is no longer needed. */
	saved(): void {
		void this.store?.delete(this.#name()).catch(() => {});
	}
	/**
	 * A file has finished opening and has its name: look for a kept copy of it. Called by the viewer
	 * once the name is known (the controller reports the document before that).
	 */
	opened(): void {
		const state = this.controller.state;
		if (!state.document || state.loading || !state.edit.sourceAvailable || state.edit.dirty) return;
		void this.#look();
	}
	async #look(): Promise<void> {
		const store = this.store;
		if (!store) return;
		const name = this.#name();
		const lookup = ++this.#lookup;
		const source = this.controller.sourceGeneration;
		const revision = this.controller.documentGeneration;
		const current = () => {
			const state = this.controller.state;
			return (
				lookup === this.#lookup &&
				source === this.controller.sourceGeneration &&
				revision === this.controller.documentGeneration &&
				name === this.#name() &&
				!state.loading &&
				!state.edit.busy &&
				!state.edit.dirty &&
				state.edit.sourceAvailable
			);
		};
		let entry: RecoveryEntry | undefined;
		try {
			entry = await store.get(name);
		} catch {
			return;
		}
		if (!entry || !current()) return;
		// A copy identical to what just opened has nothing to give back (it is also what a restore
		// opens), so it is neither offered nor removed: it stays until a copy is saved.
		if (sameBytes(entry.bytes, this.controller.exportVsdx().bytes)) return;
		const kept = entry;
		this.offer({
			name: kept.name,
			savedAt: kept.savedAt,
			restore: async () => {
				if (!current())
					throw new Error('The drawing changed. Reopen the file to recover its copy.');
				await this.controller.load(kept.bytes);
			},
			discard: async () => {
				await store.delete(name).catch(() => {});
			},
		});
	}
	/** Also copy when the page is hidden, the last moment a closing tab reliably gets. */
	wire(doc: Document): () => void {
		const hidden = () => {
			if (doc.visibilityState === 'hidden') void this.saveNow();
		};
		doc.addEventListener('visibilitychange', hidden);
		return () => doc.removeEventListener('visibilitychange', hidden);
	}
	dispose(): void {
		++this.#lookup;
		if (this.#timer !== undefined) this.timers.clearInterval(this.#timer);
		this.#timer = undefined;
		this.#enabled = false;
	}
}
