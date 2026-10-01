#!/usr/bin/env node
/**
 * publish-released.mjs: publish the packages of a release to npm, in dependency order.
 *
 * Two modes, mirroring `.github/workflows/release.yml`:
 *   --plan release-plan.json   publish every package the plan marks `release` (the scheduled run)
 *   --tag <npm-name>@<version> re-publish exactly that existing tag (manual dispatch with `tag`)
 *
 * Authentication is npm trusted publishing (GitHub OIDC): there is no token, and `--provenance`
 * attaches the build attestation. The publisher entry for each package on npmjs.com must name this
 * repository, the workflow `release.yml` and the environment `npm`.
 *
 * Safety checks before anything is uploaded, per package:
 *   - the manifest on disk is the version being published (a re-publish checks out the tag),
 *   - no dependency range uses the `file:` or `link:` protocol, and no `workspace:` range names
 *     anything but a sibling of this repo, because no consumer can install them. A `workspace:`
 *     range on a sibling (office-ui on ooxml-core) is published as `^<sibling version on disk>`:
 *     the manifest is rewritten only for the `npm publish` call and restored afterwards, so the
 *     repo keeps its workspace link (npm never rewrites `workspace:` itself),
 *   - a sibling the package requires is already on npm (core is published before ui),
 *   - every other range on a sibling is satisfied by the version that sibling ships with,
 *   - a version that already exists on the registry is skipped, so a re-run is idempotent,
 *   - a version older than the registry's `latest` is published under the `old` dist-tag so it
 *     cannot move `latest` backwards.
 *
 * `--manual` publishes without provenance, for the one-off first publish from a maintainer machine
 * (the manifest rewrite and every check still run).
 * `--dry-run` runs all the checks and prints the `npm publish` commands without running them.
 * Node >= 22, npm >= 11.5.1 (trusted publishing) on the publishing runner.
 */

import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
	cmpSemver,
	isWorkspaceRange,
	PACKAGES,
	presentPackages,
	satisfies,
} from './release-plan.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const packages = () => presentPackages(ROOT, PACKAGES);
const FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies', 'devDependencies'];
const npm = (args, options = {}) =>
	spawnSync('npm', args, {
		cwd: ROOT,
		encoding: 'utf8',
		shell: process.platform === 'win32',
		...options,
	});

/** `{ key, npm, dir, version }` targets, in dependency order, from a plan or a tag. */
export function resolveTargets({ plan, tag }) {
	if (tag) {
		if (/[^A-Za-z0-9.@/_-]/u.test(tag) || !tag.includes('@', 1)) {
			throw new Error(`Invalid tag '${tag}'. Expected <npm-name>@<semver>.`);
		}
		const at = tag.lastIndexOf('@');
		const [name, version] = [tag.slice(0, at), tag.slice(at + 1)];
		const key = Object.keys(packages()).find((k) => PACKAGES[k].npm === name);
		if (!key) throw new Error(`Unknown package '${name}' (from tag ${tag}).`);
		return [{ key, npm: name, dir: PACKAGES[key].dir, version }];
	}
	return plan.order
		.filter((key) => plan.packages[key].release)
		.map((key) => ({ key, ...pick(plan.packages[key]) }));
}
const pick = ({ npm: name, dir, version }) => ({ npm: name, dir, version });

/**
 * The manifest as it must be published: `workspace:` ranges on siblings become `^<version>`.
 * Throws when the manifest on disk is not publishable as `target`.
 */
export function publishManifest(target, versions = workspaceVersions(), root = ROOT) {
	const manifest = JSON.parse(readFileSync(join(root, target.dir, 'package.json'), 'utf8'));
	if (manifest.name !== target.npm)
		throw new Error(`${target.dir} is ${manifest.name}, not ${target.npm}.`);
	if (manifest.version !== target.version) {
		throw new Error(`${target.npm} is ${manifest.version} on disk, release is ${target.version}.`);
	}
	if (manifest.private) throw new Error(`${target.npm} is private and cannot be published.`);
	for (const field of FIELDS) {
		for (const [dep, range] of Object.entries(manifest[field] ?? {})) {
			const sibling = versions.get(dep);
			if (isWorkspaceRange(range) && sibling) {
				manifest[field][dep] = `^${sibling}`;
			} else if (/^(?:workspace|file|link):/u.test(range)) {
				throw new Error(
					`${target.npm}: ${field}["${dep}"] is "${range}", which cannot be installed.`,
				);
			} else if (field !== 'devDependencies' && sibling && !satisfies(range, sibling)) {
				throw new Error(`${target.npm}: ${dep} is "${range}" but that package is at ${sibling}.`);
			}
		}
	}
	return manifest;
}

/** Throws when the manifest on disk is not publishable as `target`. */
export const verifyManifest = (target, versions, root) =>
	void publishManifest(target, versions, root);

function workspaceVersions() {
	return new Map(
		Object.values(packages()).map((meta) => {
			const manifest = JSON.parse(readFileSync(join(ROOT, meta.dir, 'package.json'), 'utf8'));
			return [meta.npm, manifest.version];
		}),
	);
}

/** Registry state of `name@version`: 'exists', 'missing', or throws on an unexpected error. */
function registryState(name, version) {
	const result = npm(['view', `${name}@${version}`, 'version']);
	if (result.status === 0 && result.stdout.trim()) return 'exists';
	if (/E404|404 Not Found|is not in this registry/iu.test(`${result.stdout}${result.stderr}`)) {
		return 'missing';
	}
	throw new Error(`Could not query npm for ${name}@${version}: ${result.stderr || result.stdout}`);
}

function distTag(name, version) {
	const latest = npm(['view', name, 'version']);
	const current = latest.status === 0 ? latest.stdout.trim() : '';
	return /^\d+\.\d+\.\d+$/u.test(current) && cmpSemver(version, current) < 0 ? 'old' : 'latest';
}

function main() {
	const argv = process.argv.slice(2);
	const value = (flag) => argv[argv.indexOf(flag) + 1];
	const dryRun = argv.includes('--dry-run');
	const targets = resolveTargets(
		argv.includes('--tag')
			? { tag: value('--tag') }
			: {
					plan: JSON.parse(
						readFileSync(value('--plan') ?? join(ROOT, 'release-plan.json'), 'utf8'),
					),
				},
	);
	if (targets.length === 0) {
		console.log('Nothing to publish.');
		return;
	}
	for (const target of targets) {
		console.log(`--- ${target.npm}@${target.version} ---`);
		const manifest = publishManifest(target);
		if (registryState(target.npm, target.version) === 'exists') {
			console.log('already published, skipping');
			continue;
		}
		// A sibling this package requires must already be installable (core is published first).
		for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
			for (const [dep, range] of Object.entries(manifest[field] ?? {})) {
				const version = range.replace(/^[\^~]/u, '');
				if (!dryRun && workspaceVersions().has(dep) && registryState(dep, version) !== 'exists') {
					throw new Error(`${target.npm} needs ${dep}@${version}, which is not on npm yet.`);
				}
			}
		}
		// `--manual` is for the one-off first publish from a maintainer machine: provenance needs the
		// OIDC token of a CI run, so it is omitted. Every other check and rewrite is unchanged.
		const args = [
			'publish',
			...(argv.includes('--manual') ? [] : ['--provenance']),
			'--access',
			'public',
			'--tag',
			distTag(target.npm, target.version),
		];
		if (dryRun) {
			console.log(`dry run: (cd ${target.dir} && npm ${args.join(' ')})`);
			continue;
		}
		// Publish from the package's own directory with the rewritten manifest, then restore it.
		const manifestPath = join(ROOT, target.dir, 'package.json');
		const original = readFileSync(manifestPath, 'utf8');
		let result;
		try {
			writeFileSync(manifestPath, `${JSON.stringify(manifest, null, '\t')}\n`);
			result = npm(args, { cwd: join(ROOT, target.dir), stdio: 'inherit' });
		} finally {
			writeFileSync(manifestPath, original);
		}
		if (result.status !== 0) throw new Error(`Publishing ${target.npm}@${target.version} failed.`);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
