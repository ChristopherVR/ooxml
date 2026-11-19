#!/usr/bin/env node
/**
 * Attaches the tarballs and sigstore bundles `scripts/sign-release-assets.mjs` made to their GitHub
 * releases. It runs in a job that can write releases, so it runs no third-party code, and it treats
 * the signing job's files as data: a tarball is uploaded only when its sha512 is the integrity the
 * npm registry records for that version.
 *
 * Usage:
 *   node scripts/attach-release-assets.mjs --plan release-plan.json [--dir <assets>]
 *   node scripts/attach-release-assets.mjs --tag <npm-name>@<semver> [--dir <assets>]
 *   --dry-run   print what would be uploaded
 */

import { createHash } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { bundleName, parseArgs, run, tagOf, tarballName, targetsOf } from './release-assets.mjs';

/** The Subresource Integrity string npm records for a tarball. */
export const integrityOf = (bytes) =>
	`sha512-${createHash('sha512').update(bytes).digest('base64')}`;

function attachTarget(target, directory, dryRun) {
	const tag = tagOf(target);
	const tarball = join(directory, tarballName(target));
	const bundle = join(directory, bundleName(target));
	if (run('gh', ['release', 'view', tag]).status !== 0) {
		console.warn(`::warning::No GitHub release ${tag}; nothing to attach to.`);
		return true;
	}
	if (!existsSync(tarball) || !existsSync(bundle)) {
		console.warn(`::warning::${tag} was not signed; nothing to attach.`);
		return true;
	}
	const registry = run('npm', ['view', tag, 'dist.integrity']);
	if (registry.status !== 0 || registry.stdout.trim() !== integrityOf(readFileSync(tarball))) {
		console.error(`::error::${tag}: the signed tarball is not the one the npm registry serves.`);
		return false;
	}
	if (dryRun) {
		console.log(`[dry-run] upload ${tarball} and ${bundle} to ${tag}`);
		return true;
	}
	const result = run('gh', ['release', 'upload', tag, tarball, bundle, '--clobber']);
	if (result.status !== 0) {
		console.error(`::error::${tag}: 'gh release upload' failed:\n${result.stderr}`);
		return false;
	}
	console.log(`Attached ${tarballName(target)} and ${bundleName(target)} to ${tag}`);
	return true;
}

function main() {
	const options = parseArgs(process.argv.slice(2));
	const failed = targetsOf(options).filter(
		(target) => !attachTarget(target, options.dir, options.dryRun),
	);
	if (failed.length > 0) {
		console.error(`Failed to attach: ${failed.map(tagOf).join(', ')}`);
		process.exit(1);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
