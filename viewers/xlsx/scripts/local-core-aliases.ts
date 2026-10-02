import { existsSync } from 'node:fs';
import { resolve } from 'node:path';

/**
 * Opt-in source aliases for developing against an unreleased `ooxml` checkout. Set
 * `OOXML_CORE_SRC` to that checkout (for example `../ooxml`) and Vite and Vitest resolve
 * `ooxml-core/<area>` and `ooxml-ui/<entry>` to its TypeScript sources, so core edits show up
 * without rebuilding or reinstalling. Unset (CI, releases), the published packages are used.
 */
export function localCoreAliases(root: string): { find: RegExp; replacement: string }[] {
	const checkout = process.env['OOXML_CORE_SRC'];
	if (!checkout) return [];
	const base = resolve(root, checkout);
	if (!existsSync(resolve(base, 'src/xlsx/index.ts')))
		throw new Error(`OOXML_CORE_SRC=${checkout} has no src/xlsx/index.ts`);
	return [
		{ find: /^ooxml-core\/xlsx\/load$/, replacement: resolve(base, 'src/xlsx/load/index.ts') },
		{ find: /^ooxml-core\/([a-z]+)$/, replacement: resolve(base, 'src/$1/index.ts') },
		{ find: /^ooxml-core$/, replacement: resolve(base, 'src/index.ts') },
		{ find: /^ooxml-ui\/([a-z]+)$/, replacement: resolve(base, 'packages/ui/src/$1.ts') },
		{ find: /^ooxml-ui$/, replacement: resolve(base, 'packages/ui/src/index.ts') },
	];
}
