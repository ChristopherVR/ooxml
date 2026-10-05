/**
 * Merges the string files of one locale folder (`shell.ts`, `grid.ts`, `commands.ts`, ...). Each
 * file may export its table under any name (or as default); every exported plain object whose
 * values are all strings is merged, in file-name order, so agents can add files independently.
 */
export function mergeStringModules(
	modules: Record<string, Record<string, unknown>>,
): Readonly<Record<string, string>> {
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
