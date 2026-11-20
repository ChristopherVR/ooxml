// Which workspaces a change to `bun.lock`'s resolved packages reaches (ci-manifests.mjs). Bun's
// isolated linker gives a workspace only what its own dependencies resolve to, so a package that
// moved reaches the workspaces whose dependency tree, in the lock before or after, holds it.

/** A lockfile key's package names, outermost first: `a/@s/b/c` is `a`, `@s/b`, `c`. */
const segments = (key) => {
	const parts = key.split('/');
	const out = [];
	for (let i = 0; i < parts.length; i++)
		out.push(parts[i].startsWith('@') ? `${parts[i]}/${parts[++i]}` : parts[i]);
	return out;
};
const DEPENDENCY_FIELDS = [
	'dependencies',
	'devDependencies',
	'optionalDependencies',
	'peerDependencies',
];
const names = (meta, fields) => fields.flatMap((field) => Object.keys(meta?.[field] ?? {}));

/** The keys of `packages` whose entry differs between two parsed locks, or that only one has. */
export function changedPackages(a, b) {
	const pa = a.packages ?? {};
	const pb = b.packages ?? {};
	return new Set(
		[...new Set([...Object.keys(pa), ...Object.keys(pb)])].filter(
			(key) => JSON.stringify(pa[key]) !== JSON.stringify(pb[key]),
		),
	);
}

/** Every package key each workspace's dependencies resolve to in a parsed lock, transitively. */
export function workspaceClosures(lock) {
	const packages = lock.packages ?? {};
	const workspaces = lock.workspaces ?? {};
	// A dependency `name` of key `from` resolves to the nearest `<prefix>/<name>` going outwards.
	const resolve = (from, name) => {
		const outer = segments(from);
		for (let depth = outer.length; depth >= 0; depth--) {
			const key = [...outer.slice(0, depth), name].join('/');
			if (key in packages) return key;
		}
		return undefined;
	};
	const direct = (path) => {
		const workspace = workspaces[path] ?? {};
		const from = path === '' ? '' : (workspace.name ?? '');
		return names(workspace, DEPENDENCY_FIELDS)
			.map((name) => (from ? resolve(from, name) : name in packages ? name : undefined))
			.filter(Boolean);
	};
	const closures = new Map();
	for (const path of Object.keys(workspaces)) {
		const seen = new Set();
		const stack = direct(path);
		while (stack.length) {
			const key = stack.pop();
			if (seen.has(key)) continue;
			seen.add(key);
			const [id, , meta] = packages[key] ?? [];
			const workspace = /@workspace:(.*)$/.exec(id ?? '')?.[1];
			if (workspace !== undefined) stack.push(...direct(workspace));
			else
				for (const name of names(meta, [
					'dependencies',
					'optionalDependencies',
					'peerDependencies',
				])) {
					const dep = resolve(key, name);
					if (dep) stack.push(dep);
				}
		}
		closures.set(path, seen);
	}
	return closures;
}

/**
 * The workspace paths whose dependency trees, before or after, hold a changed package; undefined
 * when a changed package is in no workspace's tree (the lock cannot say who installs it).
 */
export function workspacesReaching(a, b) {
	const changed = changedPackages(a, b);
	const reached = new Set();
	const placed = new Set();
	for (const lock of [a, b])
		for (const [path, closure] of workspaceClosures(lock))
			for (const key of changed)
				if (closure.has(key)) {
					reached.add(path);
					placed.add(key);
				}
	if (placed.size < changed.size) return undefined;
	return [...reached].sort();
}
