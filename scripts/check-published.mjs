#!/usr/bin/env node
/**
 * check-published.mjs: fail when a publishable package's version on disk is not on npm.
 *
 * A release that bumps and tags a package but dies before `npm publish` leaves it stranded: the
 * next planner run sees the tag and skips it, so nothing ever retries. This check runs at the end
 * of every release run (including the hourly no-op ones), waits for the registry to serve versions
 * published a moment ago, and prints the exact dispatch that publishes each missing package.
 *
 *   node scripts/check-published.mjs [--wait-minutes 15]
 */

import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { PACKAGES, presentPackages } from './release-plan.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** `{ name, version, dir }` of every non-private package present in this checkout. */
export function publishablePackages(root = ROOT) {
	return Object.values(presentPackages(root, PACKAGES)).flatMap((meta) => {
		const manifest = JSON.parse(readFileSync(join(root, meta.dir, 'package.json'), 'utf8'));
		return manifest.private
			? []
			: [{ name: manifest.name, version: manifest.version, dir: meta.dir }];
	});
}

/** The packages whose version `exists(name, version)` says is not on the registry. */
export const findUnpublished = (packages, exists) =>
	packages.filter((pkg) => !exists(pkg.name, pkg.version));

/** The command that publishes one stranded package (documented in docs/releasing.md). */
export const recoveryCommand = (pkg) =>
	`gh workflow run release.yml -f tag=${pkg.name}@${pkg.version}`;

function existsOnNpm(name, version) {
	const result = spawnSync('npm', ['view', `${name}@${version}`, 'version'], {
		encoding: 'utf8',
		shell: process.platform === 'win32',
	});
	if (result.status === 0 && result.stdout.trim()) return true;
	if (/E404|404 Not Found|is not in this registry/iu.test(`${result.stdout}${result.stderr}`)) {
		return false;
	}
	throw new Error(`Could not query npm for ${name}@${version}: ${result.stderr || result.stdout}`);
}

function main() {
	const argv = process.argv.slice(2);
	const waitMinutes = argv.includes('--wait-minutes')
		? Number(argv[argv.indexOf('--wait-minutes') + 1])
		: 0;
	const deadline = Date.now() + waitMinutes * 60_000;
	const packages = publishablePackages();
	let missing = findUnpublished(packages, existsOnNpm);
	// The registry can take minutes to serve a tarball it has just accepted.
	while (missing.length > 0 && Date.now() < deadline) {
		console.log(`waiting for ${missing.map((pkg) => `${pkg.name}@${pkg.version}`).join(', ')}`);
		Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 30_000);
		missing = findUnpublished(missing, existsOnNpm);
	}
	if (missing.length === 0) {
		console.log(`All ${packages.length} publishable packages are on npm at their on-disk version.`);
		return;
	}
	for (const pkg of missing) {
		console.error(
			`::error::${pkg.name}@${pkg.version} is not on npm. Publish it with: ${recoveryCommand(pkg)}`,
		);
	}
	process.exit(1);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
