/** In-memory Web Storage for the assistant's Node tests; installs both browser stores. */
class MemoryStorage {
	#items = new Map();
	get length() {
		return this.#items.size;
	}
	key(index) {
		return [...this.#items.keys()][index] ?? null;
	}
	getItem(name) {
		return this.#items.get(name) ?? null;
	}
	setItem(name, value) {
		this.#items.set(name, String(value));
	}
	removeItem(name) {
		this.#items.delete(name);
	}
	clear() {
		this.#items.clear();
	}
	/** Every stored value joined, to assert a secret appears nowhere. */
	dump() {
		return [...this.#items].join('\n');
	}
}

export function installStorage() {
	globalThis.localStorage = new MemoryStorage();
	globalThis.sessionStorage = new MemoryStorage();
}
