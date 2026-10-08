/** A flat dictionary keyed by an English phrase or a dotted id (`ribbon.home.paste`). */
export type MessageDictionary = Readonly<Record<string, string>>;

/** Every locale's dictionary, keyed by its code. */
export type LocaleDictionaries<L extends string = string> = Readonly<Record<L, MessageDictionary>>;

/**
 * Merges the string files of one locale folder (`shell.ts`, `grid.ts`, ...). Each module may export
 * its table under any name (or as default); every exported plain object whose values are all
 * strings is merged, in module-path order, so files can be added independently. Later keys win.
 */
export function mergeMessageModules(
	modules: Record<string, Record<string, unknown>>,
): MessageDictionary {
	const merged: Record<string, string> = {};
	for (const path of Object.keys(modules).sort()) {
		const module = modules[path];
		if (!module) continue;
		for (const value of Object.values(module)) {
			if (!value || typeof value !== 'object' || Array.isArray(value)) continue;
			const entries = Object.entries(value as Record<string, unknown>);
			if (!entries.every(([, text]) => typeof text === 'string')) continue;
			for (const [key, text] of entries) merged[key] = text as string;
		}
	}
	return merged;
}

/** Keys of `baseline` that `dictionary` does not define (for locale completeness tests). */
export function missingMessageKeys(
	baseline: MessageDictionary,
	dictionary: MessageDictionary,
): string[] {
	return Object.keys(baseline).filter((key) => !(key in dictionary));
}
