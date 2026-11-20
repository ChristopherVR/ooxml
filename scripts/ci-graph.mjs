// The import graph the CI planner reads to decide which viewers a change to the shared core or UI
// can reach (scripts/ci-plan.mjs). A source file reaches a viewer when a file the viewer owns
// (`viewers/<name>/`, `demos/<name>/`, `e2e/<name>/`) imports it, directly or through any chain of
// imports, package subpaths (`ooxml-core/<area>`, `ooxml-ui/<entry>`) and workspace packages
// included.
//
// It errs towards reaching more. A specifier it cannot resolve to a file stands for every file in
// the folder it points into (the whole package for an `ooxml-core` or `ooxml-ui` subpath), an import
// of another workspace package reaches every file of that package, and comments that look like
// imports count. A file the graph cannot speak for (anything but a module, the packages' build
// scripts and configuration) has no answer, and the planner falls back to its area rules.
import { execFileSync } from 'node:child_process';
import { dirname, posix } from 'node:path';

/** Files that are modules: what an import can name and what can hold one. */
export const MODULE = /\.(ts|tsx|mts|cts|js|jsx|mjs|cjs|vue|svelte|css)$/;
/** Roots whose files take part: the two packages and what the viewers own. */
const ROOTS = /^(src\/core|src\/ui|viewers|demos|e2e|apps|mcp)\//;
const EXTENSIONS = ['', '.ts', '.tsx', '.mts', '.js', '.mjs', '.jsx', '.vue', '.svelte', '.css'];
const INDEXES = ['/index.ts', '/index.tsx', '/index.mts', '/index.js', '/index.mjs'];
/** Shipped code of the core and the UI: files under the packages but outside these have no answer. */
const SPEAKS_FOR = /^(src\/core\/(?!scripts\/)[^/]+\/|src\/ui\/src\/)/;
/**
 * Who runs a viewer-owned path: `<viewer>` for its checks and browser suite, `<viewer>:mcp` for its
 * standalone MCP server and `pptx:<package>` for one PowerPoint package's unit tests.
 */
export function ownerOf(path) {
	const match = /^(viewers|demos|e2e)\/([^/]+)\//.exec(path);
	if (!match) return undefined;
	const [prefix, root, viewer] = match;
	if (root !== 'viewers') return viewer;
	const rest = path.slice(prefix.length);
	if (rest.startsWith('mcp/')) return `${viewer}:mcp`;
	const leg = viewer === 'pptx' && /^packages\/([^/]+)\//.exec(rest)?.[1];
	return leg ? `pptx:${leg}` : viewer;
}
/** Viewer files its checks never run: maintenance scripts, the docs site, the old repository's CI. */
const NOT_RUN = /^viewers\/[^/]+\/(scripts|docs|\.github)\//;
/** Files of the packages that do not ship, which only an explicit import reaches. */
const UNSHIPPED =
	/^src\/(core|ui)\/(scripts\/|.*(\.test\.[cm]?[jt]sx?$|\/(__tests__|__fixtures__|test-support)\/))/;

const PATTERNS = [
	// import x from '...', export { y } from '...', across lines
	/\b(?:import|export)\b[^'"`;]*?\bfrom\s*['"]([^'"]+)['"]/g,
	// import '...' and import('...')
	/\bimport\s*\(?\s*['"]([^'"]+)['"]/g,
	/\brequire\s*\(\s*['"]([^'"]+)['"]\s*\)/g,
	/\bnew\s+URL\s*\(\s*['"]([^'"]+)['"]\s*,\s*import\.meta\.url/g,
	/@import\s+(?:url\(\s*)?['"]([^'"]+)['"]/g,
];
/** Imports whose target is only partly written: the folder they name stands for all of it. */
const PARTIAL = [
	/\bimport\s*\(\s*`([^`$]*)/g,
	/\bimport\.meta\.glob\s*\(\s*\[?\s*['"]([^'"*{]*)/g,
	/\bnew\s+URL\s*\(\s*`([^`$]*)`?[^)]*import\.meta\.url/g,
];

/** The specifiers a file names: `exact` ones and the `partial` prefixes of dynamic ones. */
export function specifiers(source) {
	// Comments often show an import as an example. Only comments that open a line go: after code,
	// `//` and `/*` may sit inside a string (a URL, a glob).
	const text = source.replace(/^\s*\/\/.*$/gm, '').replace(/^\s*\/\*[\s\S]*?\*\//gm, '');
	const exact = new Set();
	const partial = new Set();
	for (const pattern of PATTERNS) for (const m of text.matchAll(pattern)) exact.add(m[1]);
	for (const pattern of PARTIAL) for (const m of text.matchAll(pattern)) partial.add(m[1]);
	return { exact: [...exact], partial: [...partial] };
}

/** `ooxml-ui` subpaths and the source file each tsup entry builds (`index` is the package root). */
export function uiEntries(configs) {
	const entries = new Map();
	for (const [file, text] of configs) {
		const prefix = file.endsWith('tsup.pptx.config.ts') ? 'pptx/' : '';
		for (const m of text.matchAll(/['"]?([\w/.-]+)['"]?\s*:\s*['"](src\/[^'"]+)['"]/g)) {
			const key = `${prefix}${m[1]}`.replace(/(^|\/)index$/, '');
			entries.set(key.replace(/\/$/, ''), `src/ui/${m[2]}`);
		}
	}
	return entries;
}

/**
 * The graph over `files` (path to text). `packages` maps each workspace package's name to its folder.
 * Returns `reach(file)`: the viewers that can reach a file, or undefined when the graph cannot say.
 */
export function importGraph(files, packages, uiEntry = new Map()) {
	const has = (path) => files.has(path);
	// What each file imports: files, folders (a partial or unresolved specifier) and packages.
	const imports = new Map();
	const add = (from, kind, value) => {
		const edges =
			imports.get(from) ?? imports.set(from, { files: [], folders: [], packages: [] }).get(from);
		edges[kind].push(value);
	};
	const pick = (base) => {
		const stem = base.replace(/.(m?js|jsx)$/, '');
		for (const candidate of [
			...EXTENSIONS.map((ext) => base + ext),
			...['.ts', '.tsx', '.mts'].map((ext) => stem + ext),
			...INDEXES.map((index) => base + index),
		])
			if (has(candidate)) return candidate;
		return undefined;
	};
	const folders = new Set();
	for (const path of files.keys())
		for (let end = path.lastIndexOf('/'); end > 0; end = path.lastIndexOf('/', end - 1))
			folders.add(path.slice(0, end));
	// A specifier that names a folder stands for the folder; one that names no file, for its parent.
	const near = (path) => {
		const base = path.replace(/\/+$/, '');
		return folders.has(base) ? base : posix.dirname(base);
	};
	const link = (from, spec, partial) => {
		const bare = spec.replace(/[?#].*$/, '');
		if (bare.startsWith('.')) {
			const base = posix.normalize(posix.join(posix.dirname(from), bare));
			const target = !partial && pick(base);
			if (target) add(from, 'files', target);
			else add(from, 'folders', near(base));
			return;
		}
		const scoped = bare.startsWith('@');
		const parts = bare.split('/');
		const name = parts.slice(0, scoped ? 2 : 1).join('/');
		const sub = parts.slice(scoped ? 2 : 1).join('/');
		if (name === 'ooxml-core' || name === 'ooxml-ui') {
			const entry = uiEntry.get(sub);
			const target =
				!partial &&
				(name === 'ooxml-core'
					? pick(sub ? `src/core/${sub}` : 'src/core/index')
					: (entry && has(entry) && entry) || pick(sub ? `src/ui/src/${sub}` : 'src/ui/src/index'));
			if (target) add(from, 'files', target);
			else add(from, 'folders', name === 'ooxml-core' ? 'src/core' : 'src/ui');
		} else if (packages.has(name) && name !== 'ooxml-core' && name !== 'ooxml-ui')
			add(from, 'packages', packages.get(name));
	};
	for (const [path, text] of files) {
		const { exact, partial } = specifiers(text);
		for (const spec of exact) link(path, spec, false);
		for (const spec of partial) link(path, spec, true);
	}

	// Every file under a folder, from one sorted list.
	const sorted = [...files.keys()].sort();
	const under = (folder) => {
		const prefix = `${folder}/`;
		let lo = 0;
		let hi = sorted.length;
		while (lo < hi) {
			const mid = (lo + hi) >> 1;
			if (sorted[mid] < prefix) lo = mid + 1;
			else hi = mid;
		}
		const out = [];
		for (let i = lo; i < sorted.length && sorted[i].startsWith(prefix); i++)
			if (!UNSHIPPED.test(sorted[i]) && !NOT_RUN.test(sorted[i])) out.push(sorted[i]);
		return out;
	};
	// Walk forward from each viewer's own files: what the walk meets, the viewer reaches.
	const owned = new Map();
	for (const path of sorted) {
		const key = ownerOf(path);
		if (key && !NOT_RUN.test(path)) (owned.get(key) ?? owned.set(key, []).get(key)).push(path);
	}
	const reached = new Map();
	// How the walk met each file (the importing file), to explain a reach.
	const parents = new Map();
	for (const [owner, seeds] of owned) {
		const seen = new Set(seeds);
		const parent = new Map();
		parents.set(owner, parent);
		const queue = [...seeds];
		const visit = (path, from) => {
			if (!seen.has(path)) {
				parent.set(path, from);
				seen.add(path);
				queue.push(path);
			}
		};
		const expanded = new Set();
		while (queue.length) {
			const from = queue.pop();
			const edges = imports.get(from);
			if (!edges) continue;
			for (const path of edges.files) visit(path, from);
			for (const dir of edges.packages)
				if (!expanded.has(dir)) {
					expanded.add(dir);
					for (const path of under(dir)) visit(path, from);
				}
			// A folder is a guess at a path, not an import: it stays among the files this walk's
			// owner runs and the shared ones (a config's `new URL('.')` is not its MCP server).
			for (const folder of edges.folders)
				if (!expanded.has(`folder:${folder}`)) {
					expanded.add(`folder:${folder}`);
					for (const path of under(folder))
						if ((ownerOf(path) ?? owner) === owner) visit(path, from);
				}
		}
		reached.set(owner, seen);
	}
	function reach(file) {
		if (!MODULE.test(file) || !SPEAKS_FOR.test(file)) return undefined;
		return new Set([...reached].filter(([, seen]) => seen.has(file)).map(([owner]) => owner));
	}
	/** The import chain from a file `owner` owns to `file`, outermost first. */
	reach.why = (owner, file) => {
		const chain = [];
		for (let path = file; path; path = parents.get(owner)?.get(path)) chain.unshift(path);
		return chain;
	};
	return reach;
}

/**
 * The tracked files under `roots` as path to text, read from git's index in one `cat-file` batch
 * (opening some 19,000 files one by one takes half a minute on Windows).
 */
function trackedFiles(cwd, keep) {
	const git = (args, input) => execFileSync('git', args, { cwd, input, maxBuffer: 1 << 30 });
	const entries = git(['ls-files', '-s', '-z'])
		.toString('utf8')
		.split('\0')
		.filter(Boolean)
		.map((line) => {
			const tab = line.indexOf('\t');
			return { id: line.slice(0, tab).split(' ')[1], path: line.slice(tab + 1) };
		})
		.filter(({ path }) => keep(path));
	const ids = entries.map(({ id }) => `${id}\n`).join('');
	const out = git(['cat-file', '--batch'], ids);
	const files = new Map();
	let at = 0;
	for (const { path } of entries) {
		// Each blob is `<id> blob <size>\n<content>\n`.
		const header = out.indexOf(0x0a, at);
		const size = Number(out.toString('latin1', at, header).split(' ')[2]);
		files.set(path, out.toString('utf8', header + 1, header + 1 + size));
		at = header + 1 + size + 1;
	}
	return files;
}

/** The graph of the checkout at `cwd`, from the files git tracks there. */
export function repositoryGraph(cwd = '.') {
	const tracked = trackedFiles(
		cwd,
		(path) => ROOTS.test(path) && (MODULE.test(path) || path.endsWith('/package.json')),
	);
	const files = new Map();
	const packages = new Map();
	for (const [path, text] of tracked) {
		if (MODULE.test(path)) files.set(path, text);
		else
			try {
				const name = JSON.parse(text).name;
				if (name) packages.set(name, dirname(path));
			} catch {
				// A manifest that does not parse names no package.
			}
	}
	const configs = ['src/ui/tsup.config.ts', 'src/ui/tsup.pptx.config.ts']
		.filter((path) => files.has(path))
		.map((path) => [path, files.get(path)]);
	return importGraph(files, packages, uiEntries(configs));
}
