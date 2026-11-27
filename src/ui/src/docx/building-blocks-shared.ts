import {
	glossaryEntryKey,
	parseSharedBuildingBlocks,
	sharedBuildingBlocksValue,
	type GlossaryEntry,
} from 'ooxml-core/docx';

/**
 * A host's store for building blocks shared by every document (Word's Building Blocks.dotx), such
 * as a profile's preference store. Values are JSON-compatible; nothing is stored without one.
 */
export interface WordBuildingBlocksStore {
	read(): unknown | Promise<unknown>;
	write(value: unknown): void | Promise<void>;
}

/** The shared building blocks of one editor: loaded from the host's store and written back. */
export class SharedBuildingBlocks {
	#store: WordBuildingBlocksStore | undefined;
	#entries: GlossaryEntry[] = [];
	#loading: Promise<void> | undefined;
	#generation = 0;

	readonly changed: () => void;

	constructor(changed: () => void) {
		this.changed = changed;
	}

	get store(): WordBuildingBlocksStore | undefined {
		return this.#store;
	}

	/** Uses `store` from now on and reads its entries; none clears the shared entries. */
	set store(store: WordBuildingBlocksStore | undefined) {
		if (store === this.#store) return;
		this.#store = store;
		this.#generation++;
		this.#entries = [];
		this.#loading = undefined;
		if (store) void this.load();
		else this.changed();
	}

	entries(): readonly GlossaryEntry[] {
		return this.#entries;
	}

	/** Reads the store (once per store; `force` re-reads, as the Organizer does when it opens). */
	load(force = false): Promise<void> {
		const store = this.#store;
		if (!store) return Promise.resolve();
		if (this.#loading && !force) return this.#loading;
		const generation = this.#generation;
		const loading = (async () => {
			let value: unknown;
			try {
				value = await store.read();
			} catch {
				value = undefined;
			}
			if (generation !== this.#generation) return;
			this.#entries = parseSharedBuildingBlocks(value);
			this.changed();
		})();
		this.#loading = loading;
		return loading;
	}

	/** Adds or replaces an entry (same name, gallery and category) and writes the store. */
	save(entry: GlossaryEntry): Promise<void> {
		const key = glossaryEntryKey(entry);
		return this.#write([...this.#entries.filter((item) => glossaryEntryKey(item) !== key), entry]);
	}

	remove(key: string): Promise<void> {
		return this.#write(this.#entries.filter((item) => glossaryEntryKey(item) !== key));
	}

	async #write(entries: GlossaryEntry[]): Promise<void> {
		const store = this.#store;
		if (!store) return;
		const value = sharedBuildingBlocksValue(entries);
		this.#entries = value.entries;
		this.changed();
		await store.write(structuredClone(value));
	}
}
