/**
 * The release assets shared by the two steps that put them on a GitHub release: the `sign-release`
 * job packs and signs them (`scripts/sign-release-assets.mjs`, which runs cosign and so holds no
 * write token), and the `attach-release-assets` job uploads them (`scripts/attach-release-assets.mjs`).
 */

import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { resolveTargets } from './publish-released.mjs';

/** The release tag of a target: `<npm-name>@<version>`. */
export const tagOf = (target) => `${target.npm}@${target.version}`;

/** The file `npm pack <name>@<version>` writes: `@scope/name` becomes `scope-name`. */
export const tarballName = (target) =>
	`${target.npm.replace(/^@/u, '').replace('/', '-')}-${target.version}.tgz`;

/** The sigstore bundle written next to a tarball. */
export const bundleName = (target) => `${tarballName(target)}.sigstore.json`;

export const run = (command, args) =>
	spawnSync(command, args, {
		encoding: 'utf8',
		shell: process.platform === 'win32',
		stdio: ['ignore', 'pipe', 'pipe'],
	});

/** `--plan <file>` or `--tag <tag>`, `--dir <assets>` (default `$RUNNER_TEMP/release-assets`) and `--dry-run`. */
export function parseArgs(argv) {
	const options = { dryRun: false, dir: join(process.env.RUNNER_TEMP ?? '.', 'release-assets') };
	for (let index = 0; index < argv.length; index++) {
		const argument = argv[index];
		if (argument === '--dry-run') options.dryRun = true;
		else if (argument === '--plan') options.plan = argv[++index];
		else if (argument === '--tag') options.tag = argv[++index];
		else if (argument === '--dir') options.dir = argv[++index];
		else throw new Error(`Unknown argument '${argument}'.`);
	}
	if (!options.plan && !options.tag) throw new Error('Pass --plan <file> or --tag <tag>.');
	return options;
}

/** The packages the options name, from the release plan or the re-publish tag. */
export function targetsOf(options) {
	const plan = options.plan ? JSON.parse(readFileSync(options.plan, 'utf8')) : undefined;
	return resolveTargets({ plan, tag: options.tag });
}
