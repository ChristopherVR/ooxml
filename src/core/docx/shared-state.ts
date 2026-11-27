// State that one entry point writes and another reads. Every published entry (`ooxml-core/docx`,
// `/docx/load`, `/docx/mail-merge`, `/docx/embedded`) is a bundle of its own, and the CJS build
// has no shared chunks, so a module-level WeakMap exists once per bundle: the rows a merge
// repeats would be recorded in the mail-merge copy and looked for in the writer's. These
// registries live on `globalThis` under a registered symbol instead, so every copy shares one.
// The name carries a version: change it when the shape of what is stored changes.

/** The process-wide WeakMap registered under `ooxml-core.<name>`. */
export function sharedWeakMap<K extends WeakKey, V>(name: string): WeakMap<K, V> {
	const key = Symbol.for(`ooxml-core.${name}`);
	const store = globalThis as typeof globalThis & Record<symbol, unknown>;
	const existing = store[key];
	if (existing instanceof WeakMap) return existing as WeakMap<K, V>;
	const created = new WeakMap<K, V>();
	Object.defineProperty(store, key, { value: created, configurable: true, enumerable: false });
	return created;
}
