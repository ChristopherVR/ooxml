/**
 * Shipped code opens a package only through `openZip`, which applies the size and part-name
 * limits: a raw `JSZip.loadAsync` inflates whatever an attacker's headers ask for. The check reads
 * every shipped source file; what it skips is named by location, so a shipped folder that happens
 * to be called `docs` or a file with `fixture` in its name is still read.
 * Usage: node scripts/check-zip-open.mjs [root]
 */
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SHIPPED_ROOTS = ['src/core', 'src/ui/src', 'apps', 'mcp', 'viewers', 'demos', 'site'];
const GUARD = 'src/core/opc/zip-limits.ts';
const GENERATED_OR_TEST_DIR =
	/^(?:node_modules|dist|coverage|__tests__|__fixtures__|test-support|test-utils)$/;
const NOT_SHIPPED_PATH = new RegExp(
	'^(?:' +
		[
			'src/core/scripts/',
			'mcp/tests?/',
			'viewers/[^/]+/(?:docs|scripts)/',
			'viewers/[^/]+/mcp/tests?/',
			'viewers/[^/]+/packages/[^/]+/(?:scripts|tests?)/',
		].join('|') +
		')',
);
const TEST_ONLY_FILE = /\.(?:test|spec|test-util)\.[^/]+$/;
/** Test helpers outside a test folder, each imported by tests alone. */
const TEST_ONLY_PATHS = new Set(['src/core/xlsx/write/mini-package-fixtures.ts']);
const SOURCE = /\.(?:[cm]?[jt]s|[jt]sx|svelte|vue)$/;

function* shippedFiles(root, directory) {
	let entries;
	try {
		entries = readdirSync(directory, { withFileTypes: true });
	} catch {
		return;
	}
	for (const entry of entries) {
		const path = join(directory, entry.name);
		const name = relative(root, path).split(sep).join('/');
		if (entry.isDirectory()) {
			if (!GENERATED_OR_TEST_DIR.test(entry.name) && !NOT_SHIPPED_PATH.test(`${name}/`))
				yield* shippedFiles(root, path);
		} else if (SOURCE.test(entry.name) && !TEST_ONLY_FILE.test(entry.name)) yield name;
	}
}

const JSZIP_BINDINGS = [
	/\bimport\s+(?!type\b)(?:\*\s+as\s+)?([\w$]+)\s+from\s+['"]jszip['"]/g,
	/\bimport\s*\{[^}]*\bdefault\s+as\s+([\w$]+)[^}]*\}\s*from\s+['"]jszip['"]/g,
	/\b(?:const|let|var)\s+([\w$]+)\s*=\s*\(?\s*(?:await\s+)?(?:import|require)\(\s*['"]jszip['"]\s*\)/g,
];
const ALIAS =
	/\b(?:const|let|var)\s+([\w$]+)(?:\s*:[^=\n]+)?\s*=\s*(?:new\s+)?([\w$]+)(?:\.default)?\s*(?:\(\s*\))?\s*(?:;|$)/gm;

/** The names in `source` that hold JSZip or an instance of it, following plain aliases. */
function jszipNames(source) {
	const names = new Set(JSZIP_BINDINGS.flatMap((re) => [...source.matchAll(re)].map((m) => m[1])));
	for (let size = -1; size !== names.size;) {
		size = names.size;
		for (const [, alias, value] of source.matchAll(ALIAS)) if (names.has(value)) names.add(alias);
	}
	return names;
}

/** Reaching a member of JSZip by a name the text does not spell out. */
function indirectAccess(names) {
	if (names.size === 0) return undefined;
	const any = [...names].map((name) => name.replace(/\$/g, '\\$')).join('|');
	return new RegExp(
		[
			`(?<![\\w$.])(?:${any})(?:\\.default)?\\s*(?:\\?\\.)?\\[`,
			`\\bReflect\\.\\w+\\(\\s*(?:${any})\\b`,
			`\\}\\s*=\\s*(?:new\\s+)?(?:${any})\\b`,
		].join('|'),
	);
}

const NAMED_OPEN = /\bloadAsync\b|load\\u0041sync|['"`]load['"`]\s*\+|\bload\$\{/;

/** Each raw or indirect package open in shipped source outside the guard, as `path:line`. */
export function findRawZipOpens(root = ROOT) {
	const found = [];
	for (const shipped of SHIPPED_ROOTS)
		for (const path of shippedFiles(root, join(root, shipped))) {
			if (path === GUARD || TEST_ONLY_PATHS.has(path)) continue;
			const source = readFileSync(join(root, path), 'utf8');
			const indirect = source.includes('jszip') ? indirectAccess(jszipNames(source)) : undefined;
			source.split('\n').forEach((line, index) => {
				if (NAMED_OPEN.test(line) || indirect?.test(line)) found.push(`${path}:${index + 1}`);
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
