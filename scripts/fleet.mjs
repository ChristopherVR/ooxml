#!/usr/bin/env node
// One report for every repository of the suite: are its `ooxml-core` / `ooxml-ui` ranges on the
// latest release, are its other dependencies behind, and is its checkout clean and in step with
// its remote?
//
//   node scripts/fleet.mjs                 report (exit 1 when an ooxml range is behind)
//   node scripts/fleet.mjs --third-party   also list other dependencies a major version behind
//   node scripts/fleet.mjs --fetch         `git fetch` each checkout first (ahead/behind is exact)
//   node scripts/fleet.mjs --update        move every behind ooxml range to the latest release
//                                          (edits manifests only; never installs or commits)
//   node scripts/fleet.mjs --only docx-viewer,xlsx-viewer
//
// Repositories are listed in `scripts/fleet.json` and are looked for next to this checkout by
// their `dir`. Where a checkout lives elsewhere, say so in the untracked `scripts/fleet.local.json`
// (`{ "pptx-viewer": "D:/Development/pptx-viewer-new" }`).
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PACKAGES, bumpManifest } from './sync-ooxml-deps.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const SECTIONS = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'];
const RANGE = /^[\^~]?(\d+)\.(\d+)\.(\d+)/;

/** How far `latest` is ahead of the version a plain range starts at: 'major', 'minor', 'patch' or null. */
export function behind(range, latest) {
	const have = RANGE.exec(range ?? '');
	const want = RANGE.exec(latest ?? '');
	if (!have || !want) return null;
	for (const [i, level] of ['major', 'minor', 'patch'].entries())
		if (Number(want[i + 1]) !== Number(have[i + 1]))
			return Number(want[i + 1]) > Number(have[i + 1]) ? level : null;
	return null;
}

/** Repository name -> checkout directory (the config's `dir` beside this checkout, then local overrides). */
export function locate(config, local, root) {
	return Object.fromEntries(
		config.repos.map(({ name, dir }) => [name, resolve(local[name] ?? resolve(root, '..', dir))]),
	);
}

const readJson = (path, fallback) =>
	existsSync(path) ? JSON.parse(readFileSync(path, 'utf8')) : fallback;

const git = (cwd, ...args) => {
	try {
		return execFileSync('git', args, {
			cwd,
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'ignore'],
		});
	} catch {
		return '';
	}
};

function manifestsOf(dir) {
	return git(dir, 'ls-files', '--', '*package.json')
		.split('\n')
		.filter((file) => file && !file.includes('node_modules/') && !file.includes('/dist/'));
}

const registry = new Map();
async function latestOf(name) {
	if (!registry.has(name))
		registry.set(
			name,
			fetch(`https://registry.npmjs.org/${encodeURIComponent(name)}/latest`)
				.then((response) => (response.ok ? response.json() : null))
				.then((body) => body?.version ?? null)
				.catch(() => null),
		);
	return registry.get(name);
}

async function inspect(name, dir, flags, latest) {
	const result = { name, dir, git: '', ooxml: [], other: [], missing: !existsSync(dir) };
	if (result.missing) return result;
	if (flags.fetch) git(dir, 'fetch', '--quiet');
	const branch = git(dir, 'branch', '--show-current').trim();
	const counts = git(dir, 'rev-list', '--left-right', '--count', 'HEAD...@{upstream}').trim();
	const [ahead = '0', behindBy = '0'] = counts.split(/\s+/);
	const dirty = git(dir, 'status', '--porcelain').split('\n').filter(Boolean).length;
	result.git = [
		branch || 'detached',
		Number(ahead) ? `ahead ${ahead}` : '',
		Number(behindBy) ? `behind ${behindBy}` : '',
		dirty ? `${dirty} uncommitted` : '',
	]
		.filter(Boolean)
		.join(', ');
	const seen = new Set();
	for (const file of manifestsOf(dir)) {
		const path = resolve(dir, file);
		const manifest = JSON.parse(readFileSync(path, 'utf8'));
		const { manifest: next, changes } = bumpManifest(manifest, latest);
		for (const change of changes) result.ooxml.push({ file, ...change });
		if (flags.update && changes.length) {
			const raw = readFileSync(path, 'utf8');
			const indent = /^\t/m.test(raw) ? '\t' : 2;
			const eol = raw.includes('\r\n') ? '\r\n' : '\n';
			writeFileSync(path, JSON.stringify(next, null, indent).replace(/\n/g, eol) + eol);
		}
		if (!flags.thirdParty) continue;
		for (const section of SECTIONS)
			for (const [dep, range] of Object.entries(manifest[section] ?? {})) {
				if (
					PACKAGES.includes(dep) ||
					typeof range !== 'string' ||
					!RANGE.test(range) ||
					range.includes('||') ||
					range.includes(' - ')
				)
					continue;
				const key = `${dep}@${range}`;
				if (seen.has(key)) continue;
				seen.add(key);
				const newest = await latestOf(dep);
				const level = behind(range, newest);
				if (level === 'major') result.other.push({ dep, range, latest: newest });
			}
	}
	return result;
}

async function main() {
	const argv = process.argv.slice(2);
	const flags = {
		thirdParty: argv.includes('--third-party'),
		fetch: argv.includes('--fetch'),
		update: argv.includes('--update'),
	};
	const onlyAt = argv.indexOf('--only');
	const only = onlyAt >= 0 ? (argv[onlyAt + 1] ?? '').split(',') : null;
	const root = resolve(here, '..');
	const config = readJson(resolve(here, 'fleet.json'), { repos: [] });
	const local = readJson(resolve(here, 'fleet.local.json'), {});
	const dirs = locate(config, local, root);
	const latest = Object.fromEntries(
		await Promise.all(PACKAGES.map(async (name) => [name, await latestOf(name)])),
	);
	console.log(`Latest: ${PACKAGES.map((name) => `${name} ${latest[name]}`).join(', ')}\n`);
	let stale = 0;
	for (const { name } of config.repos) {
		if (only && !only.includes(name)) continue;
		const result = await inspect(name, dirs[name], flags, latest);
		if (result.missing) {
			console.log(`${name}: not found at ${result.dir} (see scripts/fleet.local.json)`);
			continue;
		}
		const files = [...new Set(result.ooxml.map((change) => change.file))];
		const state = !result.ooxml.length
			? 'ooxml up to date'
			: flags.update
				? `ooxml updated in ${files.length} manifest(s); run the install and commit`
				: `ooxml BEHIND in ${files.length} manifest(s)`;
		if (result.ooxml.length && !flags.update) stale += 1;
		console.log(`${name}: ${state} [${result.git}]`);
		for (const change of result.ooxml.slice(0, 3))
			console.log(`    ${change.file}: ${change.name} ${change.from} -> ${change.to}`);
		if (result.ooxml.length > 3) console.log(`    ... and ${result.ooxml.length - 3} more`);
		for (const item of result.other)
			console.log(`    major behind: ${item.dep} ${item.range} (latest ${item.latest})`);
	}
	if (stale) process.exit(1);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) await main();
