#!/usr/bin/env node
/**
 * release-plan.mjs: decide which packages to version and publish, with an INDEPENDENT version
 * line per package.
 *
 * Adapted from ChristopherVR/pptx-viewer `scripts/release-plan.mjs` (Apache-2.0), vendored as is from
 * ChristopherVR/docx-viewer (single-package form: only the PACKAGES table differs). The planning
 * model is the same; differences: dependencies between packages are read from the manifests
 * instead of a hand-kept `triggers` list, internal dependency ranges are rewritten when a
 * dependency is released, and a package that was never published is released at its manifest
 * version instead of being bumped. Candidate for centralisation (the config block is the only
 * repo-specific part): the same file is vendored in ooxml-core.
 *
 * Every published package carries its own version and its own git tag `<npm-name>@<version>`
 * (e.g. `@christophervr/ooxml-core@0.2.0`). From the history since each package's last tag this
 * script computes which packages changed (directly, or because an internal dependency is being
 * released) and the next version for each. The bump level follows Conventional Commits: a
 * breaking change (`!` or a BREAKING CHANGE footer) bumps major, `feat` bumps minor, anything
 * else patch. Unchanged packages get no tag, no GitHub release and no npm publish.
 *
 *   node scripts/release-plan.mjs             # print the plan (dry run, writes release-plan.json)
 *   node scripts/release-plan.mjs --no-npm    # skip npm lookups (offline)
 *   node scripts/release-plan.mjs --write     # also stamp versions and internal ranges
 *
 * It is the single source of truth for `.github/workflows/release.yml`; the rich per-package
 * detail is read back from `release-plan.json`. Under GitHub Actions it appends `any_changed` to
 * `$GITHUB_OUTPUT`.
 */

import { execFileSync } from 'node:child_process';
import { appendFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Publishable packages, in the order the planner reports them. `dir` is the source directory
 * and `npm` the published name. Dependencies between these packages are NOT listed here: any
 * `dependencies` / `peerDependencies` / `optionalDependencies` entry naming another package of
 * this table is an internal dependency, and releasing it re-releases the dependent.
 * `paths` (optional) narrows what counts as a published file, for a package that is the repo
 * root: entries ending in `/` are directories, anything else a single file.
 *
 * Two packages live here. `core` is rooted at the repository root, so `paths` lists what reaches
 * its npm tarball or decides its contents: the sources, the bundler and declaration configs, the
 * manifest (see IGNORED_MANIFEST_FIELDS) and the licence files. Docs, CI, tests and the build
 * orchestration scripts (scripts/build.mjs, scripts/ensure-built.mjs) do not release anything.
 * `ui` is a Bun workspace at packages/ui and depends on core; core never depends on it. A package
 * whose manifest does not exist yet is left out of the plan (see `presentPackages`).
 */
export const PACKAGES = {
	core: {
		dir: '.',
		npm: '@christophervr/ooxml-core',
		paths: [
			'src/',
			'scripts/pptx/merge-declarations.mjs',
			'tsconfig.json',
			'tsconfig.build.json',
			'tsconfig.pptx.json',
			'tsup.config.ts',
			'tsup.pptx.config.ts',
			'tsdown.pptx.config.ts',
			'package.json',
			'LICENSE',
			'NOTICE',
			'THIRD-PARTY-LICENSES',
		],
	},
	ui: { dir: 'packages/ui', npm: '@christophervr/office-ui' },
};

/** Paths outside any package that still change every published artifact: none for one package. */
export const GLOBAL_TRIGGERS = [];

/** package.json fields whose change never alters what a consumer receives. */
const IGNORED_MANIFEST_FIELDS = ['version', 'scripts', 'devDependencies', 'workspaces'];
const DEP_FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies'];
const BUMP_RANK = { patch: 0, minor: 1, major: 2 };
const SEMVER = /^\d+\.\d+\.\d+$/u;

export function cmpSemver(a, b) {
	const pa = a.split('.').map(Number);
	const pb = b.split('.').map(Number);
	for (let i = 0; i < 3; i++) {
		const d = (pa[i] || 0) - (pb[i] || 0);
		if (d !== 0) return d > 0 ? 1 : -1;
	}
	return 0;
}

export const maxSemver = (versions) =>
	versions.reduce((best, v) => (cmpSemver(v, best) > 0 ? v : best), '0.0.0');

export function bumpVersion(version, level) {
	const [maj, min, patch] = version.split('.').map(Number);
	if (level === 'major') return `${maj + 1}.0.0`;
	if (level === 'minor') return `${maj}.${min + 1}.0`;
	return `${maj}.${min}.${patch + 1}`;
}

/** Conventional Commit bump level of one commit message. */
export function commitLevel(subject, body = '') {
	if (/^[a-z]+(?:\([^)]*\))?!:/iu.test(subject) || /(?:^|\n)BREAKING[ -]CHANGE:/u.test(body)) {
		return 'major';
	}
	return /^feat(?:\([^)]*\))?:/u.test(subject) ? 'minor' : 'patch';
}

/**
 * Whether `version` satisfies a plain `x.y.z`, `^x.y.z`, `~x.y.z` or `*` range. Anything else
 * (comparators, unions) is reported as unsatisfied, which errs towards re-releasing.
 */
export function satisfies(range, version) {
	const text = range.trim();
	if (text === '*' || text === '') return true;
	const m = /^([\^~]?)(\d+)\.(\d+)\.(\d+)$/u.exec(text);
	if (!m) return false;
	const low = [m[2], m[3], m[4]].map(Number);
	const v = version.split('.').map(Number);
	const order = cmpSemver(version, low.join('.'));
	if (order < 0) return false;
	if (m[1] === '') return order === 0;
	if (m[1] === '~') return v[0] === low[0] && v[1] === low[1];
	if (low[0] > 0) return v[0] === low[0];
	return low[1] > 0 ? v[0] === 0 && v[1] === low[1] : v[0] === 0 && v[1] === 0 && v[2] === low[2];
}

/**
 * A range that tracks the sibling in this repository: `workspace:*`, or `*` where Bun cannot link
 * the sibling (the repository root is the core and cannot be a member of its own workspace, so the
 * UI declares `*` and tsconfig paths point at the core source). Either is published as a caret
 * range on the sibling's version at publish time and is never rewritten in the repository, so
 * `bun install` never has to resolve a version that is only published later in the same release.
 */
export const isWorkspaceRange = (range) => /^(?:workspace:|\*$)/u.test(range);

/** The packages of `table` whose manifest exists in `root` (a package may not be merged yet). */
export function presentPackages(root, table) {
	return Object.fromEntries(
		Object.entries(table).filter(([, meta]) => existsSync(join(root, meta.dir, 'package.json'))),
	);
}

/** Files that ship in a published artifact (exclude tests). */
export function isPublishedFile(path) {
	if (/\.(?:test|spec)\.[cm]?[jt]sx?$/u.test(path)) return false;
	return !path.includes('/__tests__/') && !path.includes('/e2e/');
}

/** JSON text with object keys sorted, so reordering a manifest is not a change. */
const canonical = (value) =>
	JSON.stringify(value, (_key, v) =>
		v && typeof v === 'object' && !Array.isArray(v)
			? Object.fromEntries(Object.entries(v).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
			: v,
	);

/** `dir` as a prefix for a repo-relative path: '' for the root, 'packages/ui/' otherwise. */
const posix = (dir) => (dir === '.' ? '' : `${dir.replace(/\/$/u, '')}/`);

const readJson = (path) => JSON.parse(readFileSync(path, 'utf8'));
const writeJson = (path, data) => writeFileSync(path, `${JSON.stringify(data, null, '\t')}\n`);

/** `npm view <name> version`: the version, or null when the package does not exist (E404). */
export function npmVersion(name) {
	try {
		const out = execFileSync('npm', ['view', name, 'version'], {
			encoding: 'utf8',
			shell: process.platform === 'win32',
			stdio: ['ignore', 'pipe', 'pipe'],
		});
		return out.trim() || null;
	} catch (error) {
		if (/E404|404 Not Found|is not in this registry/iu.test(String(error.stderr ?? ''))) {
			return null;
		}
		throw new Error(`Could not query npm for ${name}: ${error.stderr || error.message}`);
	}
}

/** `npm view <name>@<version> gitHead`: the commit a version was published from, or null. */
export function npmGitHead(name, version) {
	try {
		const out = execFileSync('npm', ['view', `${name}@${version}`, 'gitHead'], {
			encoding: 'utf8',
			shell: process.platform === 'win32',
			stdio: ['ignore', 'pipe', 'ignore'],
		});
		return /^[0-9a-f]{7,40}$/u.test(out.trim()) ? out.trim() : null;
	} catch {
		return null;
	}
}

/**
 * Compute the release plan.
 * @param {object} options
 * @param {string} options.root repository root (a git work tree)
 * @param {Record<string, {dir: string, npm: string, paths?: string[]}>} options.packages
 * @param {string[]} [options.globalTriggers]
 * @param {(name: string) => string | null} [options.npm] registry lookup; omit for offline
 * @param {(name: string, version: string) => string | null} [options.npmHead] `gitHead` of a
 *   published version, used to baseline a package that was published by hand and never tagged
 */
export function planRelease({ root, packages: all, globalTriggers = [], npm, npmHead }) {
	const table = presentPackages(root, all);
	const git = (args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
	const manifestPath = (meta) => join(root, meta.dir, 'package.json');
	const names = new Set(Object.values(table).map((m) => m.npm));
	const scopeOf = (meta) => meta.paths ?? [meta.dir];
	const under = (file, target) =>
		file === target || file.startsWith(target.endsWith('/') ? target : `${target}/`);
	const touches = (files, targets) => files.some((f) => targets.some((t) => under(f, t)));

	/** Internal dependencies (package keys) of one package, from its manifest at HEAD. */
	const internalDeps = (meta) => {
		const manifest = readJson(manifestPath(meta));
		const found = new Set();
		for (const field of DEP_FIELDS) {
			for (const dep of Object.keys(manifest[field] ?? {})) {
				const key = Object.keys(table).find((k) => table[k].npm === dep);
				if (key) found.add(key);
			}
		}
		return [...found];
	};

	/** Keys in dependency order: every package comes after the packages it depends on. */
	const order = [];
	const visit = (key, trail = []) => {
		if (order.includes(key)) return;
		if (trail.includes(key)) throw new Error(`Dependency cycle: ${[...trail, key].join(' -> ')}`);
		for (const dep of internalDeps(table[key])) visit(dep, [...trail, key]);
		order.push(key);
	};
	Object.keys(table).forEach((key) => visit(key));

	const baselineTag = (name) => {
		const tags = git(['tag', '--list', `${name}@*`, '--sort=-version:refname'])
			.split('\n')
			.map((t) => t.trim())
			.filter((t) => t.startsWith(`${name}@`) && SEMVER.test(t.slice(name.length + 1)));
		for (const tag of tags) {
			try {
				// A tag on HEAD is a valid baseline: it yields an empty diff (no re-release).
				git(['merge-base', '--is-ancestor', tag, 'HEAD']);
				return tag;
			} catch {
				// Tag on another line of history; keep looking.
			}
		}
		return null;
	};

	/** True when `path` only differs from `base` by fields that never reach a consumer. */
	const isNoiseManifestChange = (path, base) => {
		if (!base) return false;
		try {
			const strip = (text) => {
				const data = JSON.parse(text);
				for (const field of IGNORED_MANIFEST_FIELDS) delete data[field];
				for (const field of DEP_FIELDS) {
					for (const dep of Object.keys(data[field] ?? {})) {
						if (names.has(dep)) delete data[field][dep];
					}
				}
				return canonical(data);
			};
			return strip(git(['show', `${base}:${path}`])) === strip(git(['show', `HEAD:${path}`]));
		} catch {
			return false;
		}
	};

	const isReleaseArtifact = (path, base) =>
		/(?:^|\/)CHANGELOG\.md$/u.test(path) ||
		(/(?:^|\/)package\.json$/u.test(path) && isNoiseManifestChange(path, base));

	const changedFiles = (base) =>
		(base ? git(['diff', '--name-only', `${base}..HEAD`]) : git(['ls-files']))
			.split('\n')
			.map((f) => f.trim())
			.filter((f) => f && isPublishedFile(f) && !isReleaseArtifact(f, base));

	/** Highest bump level among commits since `base` that touch published files in `scope`. */
	const bumpLevel = (base, scope) => {
		if (!base) return 'patch';
		const raw = git(['log', '--format=%H%x1f%s%x1f%b%x1e', `${base}..HEAD`, '--', ...scope]);
		let best = 'patch';
		for (const record of raw.split('\x1e')) {
			const [hash, subject = '', body = ''] = record.trim().split('\x1f');
			if (!hash) continue;
			const level = commitLevel(subject, body);
			if (BUMP_RANK[level] <= BUMP_RANK[best]) continue;
			const files = git(['show', hash, '--name-only', '--format='])
				.split('\n')
				.filter((f) => f && isPublishedFile(f) && touches([f], scope));
			if (files.length > 0) best = level;
			if (best === 'major') break;
		}
		return best;
	};

	/** The range a dependent's last release declared for `dep`, or null when it cannot be known. */
	const rangeAtBase = (meta, depMeta, base) => {
		const declared = DEP_FIELDS.map((f) => readJson(manifestPath(meta))[f]?.[depMeta.npm]).find(
			Boolean,
		);
		if (declared && !isWorkspaceRange(declared)) return declared;
		// A workspace range is published as `^<dependency version at publish time>`.
		try {
			const version = JSON.parse(
				git(['show', `${base}:${posix(depMeta.dir)}package.json`]),
			).version;
			return SEMVER.test(version) ? `^${version}` : null;
		} catch {
			return null;
		}
	};

	const plan = {};
	for (const key of order) {
		const meta = table[key];
		const published = npm ? npm(meta.npm) : null;
		const tagged = maxSemver(
			git(['tag', '--list', `${meta.npm}@*`])
				.split('\n')
				.map((t) => t.trim().slice(meta.npm.length + 1))
				.filter((v) => SEMVER.test(v)),
		);
		let base = baselineTag(meta.npm);
		// Published by hand and never tagged (the manual first publish): the registry version is the
		// baseline, not a reason to release. Prefer the commit npm recorded, else HEAD, and have the
		// workflow adopt it as the tag so the next run finds it.
		let adopt = null;
		if (!base && published !== null && SEMVER.test(published) && tagged === '0.0.0') {
			const head = npmHead?.(meta.npm, published);
			let sha = null;
			try {
				if (head) {
					git(['merge-base', '--is-ancestor', head, 'HEAD']);
					sha = git(['rev-parse', head]);
				}
			} catch {
				sha = null;
			}
			sha ??= git(['rev-parse', 'HEAD']);
			base = sha;
			adopt = {
				tag: `${meta.npm}@${published}`,
				sha,
				fromRegistryHead: sha !== git(['rev-parse', 'HEAD']),
			};
		}
		const files = changedFiles(base);
		const deps = internalDeps(meta);
		// Bump level comes from this package's own files only: a dependency's commits do not raise it.
		const scope = [...scopeOf(meta), ...globalTriggers];
		const staleDep = deps.find((d) => {
			if (!plan[d].release || !base) return false;
			if (plan[d].bump === 'major') return true;
			const range = rangeAtBase(meta, table[d], base);
			return range === null || !satisfies(range, plan[d].version);
		});
		const via = (cond, why) => (cond ? why : null);
		const reason =
			via(!base, 'no previous tag') ||
			via(touches(files, scopeOf(meta)), 'own files changed') ||
			via(staleDep, 'dependency range needs bump (major or out of range)') ||
			via(touches(files, globalTriggers), 'shared build pipeline changed');
		const release = Boolean(reason);

		const manifestVersion = readJson(manifestPath(meta)).version || '0.0.0';
		const current = maxSemver([tagged, published ?? '0.0.0', manifestVersion]);
		// Never published and never tagged: the manifest version IS the first release.
		const initial = release && !base && published === null && tagged === '0.0.0';
		const bump = !release ? null : initial ? 'initial' : bumpLevel(base, scope);
		const version = !release ? current : initial ? manifestVersion : bumpVersion(current, bump);
		plan[key] = {
			npm: meta.npm,
			dir: meta.dir,
			baseline: base,
			release,
			reason: release ? reason : null,
			initial,
			bump,
			currentVersion: current,
			version,
			tag: `${meta.npm}@${version}`,
			adopt,
			dependsOn: deps,
			manifest: `${posix(meta.dir)}package.json`,
			changelog: `${posix(meta.dir)}CHANGELOG.md`,
			includePaths: [
				...(meta.paths
					? meta.paths.map((p) => (p.endsWith('/') ? `${p}**` : p))
					: [`${meta.dir}/**`]),
				...globalTriggers,
			],
		};
	}
	return { anyChanged: Object.values(plan).some((p) => p.release), order, packages: plan };
}

/**
 * Stamp each released package's version into its package.json. A plain internal dependency range
 * of a package that is itself released is repointed at the version being released (keeping a
 * `^`/`~` prefix). `workspace:` ranges are left alone so Bun keeps linking the workspace; they
 * become a caret range on the sibling's version only in the published manifest
 * (see publish-released.mjs). A dependent that is not released keeps its range: the planner only
 * skips it while the new dependency version still satisfies that range.
 */
export function applyPlan({ root, packages: all }, plan) {
	for (const [key, meta] of Object.entries(presentPackages(root, all))) {
		const path = join(root, meta.dir, 'package.json');
		const data = readJson(path);
		const own = plan.packages[key];
		let changed = false;
		if (own.release && data.version !== own.version) {
			data.version = own.version;
			changed = true;
		}
		for (const field of own.release ? DEP_FIELDS : []) {
			for (const dep of Object.keys(data[field] ?? {})) {
				const target = Object.values(plan.packages).find((p) => p.npm === dep);
				if (!target || isWorkspaceRange(data[field][dep])) continue;
				const prefix = /^[^~]/u.exec(data[field][dep])?.[0] ?? '';
				if (data[field][dep] !== `${prefix}${target.version}`) {
					data[field][dep] = `${prefix}${target.version}`;
					changed = true;
				}
			}
		}
		if (changed) writeJson(path, data);
	}
}

function main() {
	const argv = process.argv.slice(2);
	const root = join(dirname(fileURLToPath(import.meta.url)), '..');
	const config = { root, packages: PACKAGES, globalTriggers: GLOBAL_TRIGGERS };
	const plan = planRelease({ ...config, npm: argv.includes('--no-npm') ? undefined : npmVersion });
	writeJson(join(root, 'release-plan.json'), plan);
	if (argv.includes('--write')) applyPlan(config, plan);

	console.log('Release plan (independent per-package versions):');
	for (const key of plan.order) {
		const p = plan.packages[key];
		const line = p.release
			? `${p.currentVersion} -> ${p.version} (${p.bump}; ${p.reason})  tag ${p.tag}`
			: `${p.currentVersion} (skip)`;
		console.log(`  ${key.padEnd(14)} ${line}`);
	}
	if (argv.includes('--write')) console.log('(wrote versions to released package.json files)');
	if (process.env.GITHUB_OUTPUT) {
		appendFileSync(process.env.GITHUB_OUTPUT, `any_changed=${plan.anyChanged}\n`);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
