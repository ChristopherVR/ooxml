// What a change to `bun.lock` or `src/core/package.json` really reaches. The planner used to run
// everything for either file; most of those changes are inert (a version bump), only ship a new
// `exports` entry, or only move one workspace's dependency ranges. Every pure function here takes
// strings, so the decisions are testable without git. Anything that cannot be classified is left
// as the original file, which the planner treats as "run everything".
import { execFileSync } from 'node:child_process';

/** Path the planner treats as a change to the core's shipped code only (build, package checks, tests). */
export const CORE_SHIPPING = 'src/core/package.json#shipping';
/** The "area" of that path: it belongs to no viewer, so no viewer check is planned for it. */
export const SHIPPING_AREA = 'package.json#shipping';
const CORE_MANIFEST = 'src/core/package.json';
const LOCK = 'bun.lock';
const SHIPPING_FIELDS = new Set(['exports', 'files', 'typesVersions', 'sideEffects', 'bin']);
/** Changes that alter what is installed or resolved: every framework must be checked against them. */
const DEPENDENCY_FIELDS = new Set([
	'dependencies',
	'devDependencies',
	'peerDependencies',
	'optionalDependencies',
	'peerDependenciesMeta',
	'overrides',
	'resolutions',
	'workspaces',
	'packageManager',
]);
/** Workspace folders the planner knows how to place (a viewer, a demo, a browser test package, the UI). */
const PLACEABLE_WORKSPACE = /^((viewers|demos|e2e)\/[^/]+(\/|$)|src\/ui$)/;

/** Parses a JSONC-ish lockfile: trailing commas (outside strings) are removed. Throws on bad input. */
export function parseLock(text) {
	let out = '';
	let inString = false;
	for (let i = 0; i < text.length; i++) {
		const ch = text[i];
		if (inString) {
			out += ch;
			if (ch === '\\') out += text[++i] ?? '';
			else if (ch === '"') inString = false;
		} else if (ch === '"') {
			inString = true;
			out += ch;
		} else if (ch === ',') {
			let j = i + 1;
			while (j < text.length && /\s/.test(text[j])) j++;
			if (text[j] !== '}' && text[j] !== ']') out += ch;
		} else out += ch;
	}
	return JSON.parse(out);
}

function canonical(value) {
	if (Array.isArray(value)) return value.map(canonical);
	if (value && typeof value === 'object')
		return Object.fromEntries(
			Object.keys(value)
				.sort()
				.map((key) => [key, canonical(value[key])]),
		);
	return value;
}
const same = (a, b) => JSON.stringify(canonical(a)) === JSON.stringify(canonical(b));

/**
 * Classifies a lockfile change. Returns `{ workspaces }` (the changed workspace paths, only when
 * every one can be placed) or `{ everything: true }`.
 */
export function classifyLock(before, after) {
	const bail = { everything: true };
	if (before === undefined || after === undefined) return bail;
	let a;
	let b;
	try {
		a = parseLock(before);
		b = parseLock(after);
	} catch {
		return bail;
	}
	// The resolved versions (`packages`), the lockfile version, overrides and the rest.
	for (const key of new Set([...Object.keys(a), ...Object.keys(b)]))
		if (key !== 'workspaces' && !same(a[key], b[key])) return bail;
	const wa = a.workspaces ?? {};
	const wb = b.workspaces ?? {};
	const changed = [...new Set([...Object.keys(wa), ...Object.keys(wb)])].filter(
		(key) => !same(wa[key], wb[key]),
	);
	if (changed.some((path) => !PLACEABLE_WORKSPACE.test(path))) return bail;
	return { workspaces: changed.sort() };
}

/** Classifies a change to `src/core/package.json`: `inert`, `shipping` or `everything` (and whether `deps`). */
export function classifyCoreManifest(before, after) {
	const bail = { kind: 'everything', deps: true };
	if (before === undefined || after === undefined) return bail;
	let a;
	let b;
	try {
		a = JSON.parse(before);
		b = JSON.parse(after);
	} catch {
		return bail;
	}
	const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].filter(
		(key) => key !== 'version' && !same(a[key], b[key]),
	);
	if (keys.length === 0) return { kind: 'inert', deps: false };
	if (keys.every((key) => SHIPPING_FIELDS.has(key))) return { kind: 'shipping', deps: false };
	return { kind: 'everything', deps: keys.some((key) => DEPENDENCY_FIELDS.has(key)) };
}

/**
 * Replaces `bun.lock` and `src/core/package.json` in a changed-file list by what they reach.
 * `read(path, side)` returns the file's text on the base side or the head side (undefined when
 * unreadable). Returns the new list and `deps`: true when a real dependency change is involved
 * (the resolved versions moved, or the root or core dependency fields did).
 */
export function refineManifests(files, read) {
	const out = new Set(files);
	let deps = false;
	if (out.has(LOCK)) {
		const result = classifyLock(read(LOCK, 'base'), read(LOCK, 'head'));
		if (result.everything) deps = true;
		else {
			out.delete(LOCK);
			for (const workspace of result.workspaces) out.add(`${workspace}/package.json`);
		}
	}
	if (out.has(CORE_MANIFEST)) {
		const result = classifyCoreManifest(read(CORE_MANIFEST, 'base'), read(CORE_MANIFEST, 'head'));
		if (result.kind !== 'everything') out.delete(CORE_MANIFEST);
		if (result.kind === 'shipping') out.add(CORE_SHIPPING);
		if (result.deps) deps = true;
	}
	// The root manifest is not classified: treat a change to it as a dependency change.
	if (out.has('package.json')) deps = true;
	return { files: [...out], deps };
}

const gitRun = (...args) => execFileSync('git', args, { encoding: 'utf8' });

/** The `read` function for `refineManifests` over git: the base side is the merge base with HEAD. */
export function gitReader(base, run = gitRun) {
	let from = base;
	try {
		from = run('merge-base', base, 'HEAD').trim() || base;
	} catch {
		// Keep the base itself.
	}
	return (path, side) => {
		try {
			return run('show', `${side === 'base' ? from : 'HEAD'}:${path}`);
		} catch {
			return undefined;
		}
	};
}
