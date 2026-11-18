/**
 * Shipped code opens a package only through `openZip`, which applies the size and part-name
 * limits: a raw `JSZip.loadAsync` inflates whatever an attacker's headers ask for.
 * Usage: node scripts/check-zip-open.mjs [root]
 */
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SHIPPED_ROOTS = ['src/core', 'src/ui/src', 'apps', 'mcp', 'viewers'];
const GUARD = 'src/core/opc/zip-limits.ts';
const NOT_SHIPPED_DIR =
	/^(?:node_modules|dist|scripts|docs|e2e|tests?|__tests__|__fixtures__|test-support)$/;
const NOT_SHIPPED_FILE = /\.(?:test|spec|test-util)\.|fixture/i;
const SOURCE = /\.(?:[cm]?[jt]s|[jt]sx|svelte|vue)$/;

function* shippedFiles(directory) {
	let entries;
	try {
		entries = readdirSync(directory, { withFileTypes: true });
	} catch {
		return;
	}
	for (const entry of entries) {
		const path = join(directory, entry.name);
		if (entry.isDirectory()) {
			if (!NOT_SHIPPED_DIR.test(entry.name)) yield* shippedFiles(path);
		} else if (SOURCE.test(entry.name) && !NOT_SHIPPED_FILE.test(entry.name)) yield path;
	}
}

/** Each `loadAsync` in shipped source outside the guard, as `path:line`. */
export function findRawZipOpens(root = ROOT) {
	const found = [];
	for (const shipped of SHIPPED_ROOTS)
		for (const file of shippedFiles(join(root, shipped))) {
			const path = relative(root, file).split(sep).join('/');
			if (path === GUARD) continue;
			readFileSync(file, 'utf8')
				.split('\n')
				.forEach((line, index) => {
					if (/\bloadAsync\b/.test(line)) found.push(`${path}:${index + 1}`);
				});
		}
	return found;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
	const found = findRawZipOpens(process.argv[2] ? resolve(process.argv[2]) : ROOT);
	for (const place of found)
		console.error(`${place}: open packages with openZip from opc/zip-limits, not loadAsync`);
	if (found.length > 0) process.exit(1);
}
