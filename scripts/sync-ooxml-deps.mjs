#!/usr/bin/env node
// Keeps a repository's `ooxml-core` and `ooxml-ui` ranges on the latest published versions.
//
// Every viewer repository depends on the published packages of this one. Run from a viewer's
// root (by the reusable `sync-ooxml.yml` workflow, or by hand) it rewrites each manifest's caret or
// tilde range to the latest release and reports what changed. Ranges that are not plain semver
// (`*`, `latest`, `file:`, `workspace:`) are left alone, and a range is never lowered.
//
//   node sync-ooxml-deps.mjs            rewrite manifests, print a JSON summary
//   node sync-ooxml-deps.mjs --check    change nothing; exit 1 when a manifest is behind
//
// With GITHUB_OUTPUT set, also writes `changed` (true/false) and `summary` (one line).
import { execFileSync } from 'node:child_process';
import { appendFileSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const PACKAGES = ['ooxml-core', 'ooxml-ui'];
const SECTIONS = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'];
const RANGE = /^([\^~])(\d+)\.(\d+)\.(\d+)$/;

const parts = (version) => version.split('.').map(Number);

/** True when `a` is a higher x.y.z than `b`. */
export function isNewer(a, b) {
	const [x, y] = [parts(a), parts(b)];
	for (let i = 0; i < 3; i += 1) if (x[i] !== y[i]) return (x[i] ?? 0) > (y[i] ?? 0);
	return false;
}

/**
 * Returns the manifest with every plain range of a known package moved to its latest version,
 * and the list of changes. `latest` maps a package name to its newest version.
 */
export function bumpManifest(manifest, latest) {
	const changes = [];
	const next = structuredClone(manifest);
	for (const section of SECTIONS) {
		const deps = next[section];
		if (!deps) continue;
		for (const name of Object.keys(deps)) {
			const target = latest[name];
			const match = typeof deps[name] === 'string' ? RANGE.exec(deps[name]) : null;
			if (!target || !match) continue;
			const current = `${match[2]}.${match[3]}.${match[4]}`;
			if (!isNewer(target, current)) continue;
			const range = `${match[1]}${target}`;
			changes.push({ section, name, from: deps[name], to: range });
			deps[name] = range;
		}
	}
	return { manifest: next, changes };
}

export function summarize(changes) {
	const latest = new Map();
	for (const change of changes) latest.set(change.name, change.to.replace(/^[\^~]/, ''));
	return [...latest].map(([name, version]) => `${name} ${version}`).join(' and ');
}

function trackedManifests(root) {
	const out = execFileSync('git', ['ls-files', '--', '*package.json'], {
		cwd: root,
		encoding: 'utf8',
	});
	return out
		.split('\n')
		.filter((file) => file && !file.includes('node_modules/') && !file.includes('/dist/'));
}

async function latestVersion(name) {
	const response = await fetch(`https://registry.npmjs.org/${name}/latest`);
	if (!response.ok) throw new Error(`npm registry returned ${response.status} for ${name}`);
	return (await response.json()).version;
}

async function main() {
	const check = process.argv.includes('--check');
	const root = resolve(process.cwd());
	const latest = Object.fromEntries(
		await Promise.all(PACKAGES.map(async (name) => [name, await latestVersion(name)])),
	);
	const all = [];
	for (const file of trackedManifests(root)) {
		const path = resolve(root, file);
		const raw = readFileSync(path, 'utf8');
		const { manifest, changes } = bumpManifest(JSON.parse(raw), latest);
		if (!changes.length) continue;
		all.push(...changes.map((change) => ({ file, ...change })));
		if (check) continue;
		const indent = /^\t/m.test(raw) ? '\t' : 2;
		const eol = raw.includes('\r\n') ? '\r\n' : '\n';
		writeFileSync(path, JSON.stringify(manifest, null, indent).replace(/\n/g, eol) + eol);
	}
	const summary = summarize(all);
	console.log(JSON.stringify({ latest, changed: all.length > 0, summary, changes: all }, null, 2));
	if (process.env.GITHUB_OUTPUT)
		appendFileSync(process.env.GITHUB_OUTPUT, `changed=${all.length > 0}\nsummary=${summary}\n`);
	if (check && all.length) process.exit(1);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) await main();
