#!/usr/bin/env node
/**
 * Signs every published package's tarball with a keyless sigstore bundle, for its GitHub release
 * (OpenSSF Scorecard "Signed-Releases"). `scripts/attach-release-assets.mjs` uploads them.
 *
 * Usage:
 *   node scripts/sign-release-assets.mjs --plan release-plan.json [--dir <assets>]
 *   node scripts/sign-release-assets.mjs --tag <npm-name>@<semver> [--dir <assets>]
 *   --dry-run   print what would be packed and signed
 *
 * The tarball is fetched back from the npm registry (`npm pack <name>@<version>`), so the signed
 * bytes are the bytes consumers install. Signing is keyless: `cosign sign-blob` uses the job's
 * GitHub OIDC identity (`id-token: write`), and the bundle records the workflow that signed.
 * cosign is third-party code, so this step gets no token; the attach step skips a package whose
 * GitHub release does not exist.
 */

import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { bundleName, parseArgs, run, tagOf, tarballName, targetsOf } from './release-assets.mjs';

function signTarget(target, directory, dryRun) {
	const tag = tagOf(target);
	const tarball = join(directory, tarballName(target));
	const bundle = join(directory, bundleName(target));
	if (dryRun) {
		console.log(`[dry-run] pack ${tag}, sign ${tarball} into ${bundle}`);
		return true;
	}
	const steps = [
		['npm', ['pack', tag, '--pack-destination', directory, '--silent']],
		['cosign', ['sign-blob', '--yes', '--bundle', bundle, tarball]],
	];
	for (const [command, args] of steps) {
		const result = run(command, args);
		if (result.status !== 0) {
			console.error(`::error::${tag}: '${command} ${args[0]}' failed:\n${result.stderr}`);
			return false;
		}
	}
	console.log(`Signed ${tag}: ${tarballName(target)} and ${bundleName(target)}`);
	return true;
}

function main() {
	const options = parseArgs(process.argv.slice(2));
	mkdirSync(options.dir, { recursive: true });
	const failed = targetsOf(options).filter(
		(target) => !signTarget(target, options.dir, options.dryRun),
	);
	if (failed.length > 0) {
		console.error(`Failed to sign: ${failed.map(tagOf).join(', ')}`);
		process.exit(1);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
