import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, extname, join, resolve } from 'node:path';

/** Keep bundled TS source extensionless while making declaration trees usable in NodeNext. */
export function rewriteEsmDeclarationImports(directory) {
	let changed = 0;
	for (const entry of readdirSync(directory, { withFileTypes: true })) {
		const file = join(directory, entry.name);
		if (entry.isDirectory()) {
			changed += rewriteEsmDeclarationImports(file);
			continue;
		}
		if (!/\.d\.(?:ts|mts|cts)$/.test(entry.name)) continue;
		const source = readFileSync(file, 'utf8');
		// Skip comments and unrelated strings; only imports, exports and augmentations are paths.
		const output = source.replace(
			/\/\*[\s\S]*?\*\/|\/\/[^\r\n]*|(?:\bfrom\s+|\bimport\s*(?:\(\s*)?|\bdeclare\s+module\s+)(['"])(\.{1,2}\/[^'"]+)\1|"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'/g,
			(match, quote, specifier) => {
				if (!specifier || extname(specifier)) return match;
				const base = resolve(dirname(file), specifier);
				for (const [declaration, extension] of [
					['.d.ts', '.js'],
					['.d.mts', '.mjs'],
					['.d.cts', '.cjs'],
				]) {
					const target = existsSync(base + declaration)
						? specifier + extension
						: existsSync(join(base, 'index' + declaration))
							? specifier.replace(/\/$/, '') + '/index' + extension
							: undefined;
					if (target) return match.slice(0, -specifier.length - 1) + target + quote;
				}
				return match;
			},
		);
		if (output !== source) {
			writeFileSync(file, output);
			++changed;
		}
	}
	return changed;
}
