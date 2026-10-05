#!/usr/bin/env node
/**
 * publish-local.mjs (`bun run publish:local`): publish the OpenTeams packages from a maintainer's
 * own machine, for when the CI path (release.yml, npm trusted publishing) is not available yet,
 * typically the very first version of each package name.
 *
 *   bun run publish:local                    check, build, smoke-test, show the plan, ask for `yes`
 *   bun run publish:local -- --dry-run       the same, then `npm publish --dry-run` for each package
 *   bun run publish:local -- --yes           skip the typed confirmation
 *   bun run publish:local -- --package server --package react    only these packages
 *   bun run publish:local -- --allow-dirty   allow uncommitted changes or a branch other than main
 *   bun run publish:local -- --skip-build    reuse the dist/ folders already built
 *
 * It NEVER asks for or accepts a one-time password: there is no `--otp` and no OTP prompt. Publishing
 * uses `--auth-type=web`, so when the account needs 2FA npm prints a sign-in URL and opens your
 * browser; approve it there. If npm still insists on a code, the script stops and says so; publish with a granular access
 * token that has "bypass 2FA" enabled, or through the CI trusted-publishing path. Login uses
 * `npm login --auth-type=web` (npm opens the browser); this script never reads, writes or prints a
 * token, and never passes `--provenance` (provenance needs GitHub Actions OIDC).
 *
 * Safety checks are the ones `publish-released.mjs` runs in CI, reused from it: each manifest on
 * disk is the version the release plan computes, no range uses `workspace:` / `file:` / `link:`,
 * no dependency names a private workspace package, and versions already on npm are skipped.
 */
import { spawn, spawnSync } from 'node:child_process';
import { join } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { fileURLToPath } from 'node:url';

import {
	distTag,
	provenanceEnabled,
	publishArgs,
	REGISTRY,
	registryState,
	resolveTargets,
	ROOT,
	verifyManifest,
} from './publish-released.mjs';
import { GLOBAL_TRIGGERS, npmVersion, PACKAGES, planRelease } from './release-plan.mjs';

const FLAGS = new Set(['--yes', '--dry-run', '--allow-dirty', '--skip-build', '--help']);

/** Parses the command line. Throws on anything unknown, and on any attempt to pass an OTP. */
export function parseArgs(argv) {
	const options = {
		yes: false,
		dryRun: false,
		allowDirty: false,
		skipBuild: false,
		help: false,
		packages: [],
	};
	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i];
		if (/otp/iu.test(arg)) {
			throw new Error(
				'publish:local never accepts a one-time password. Use a granular access token with ' +
					'"bypass 2FA" enabled, or publish through the CI trusted-publishing path.',
			);
		}
		if (arg === '--package') {
			const key = argv[++i];
			if (!key || !(key in PACKAGES)) {
				throw new Error(`--package needs one of: ${Object.keys(PACKAGES).join(', ')}.`);
			}
			options.packages.push(key);
			continue;
		}
		if (!FLAGS.has(arg)) throw new Error(`Unknown argument "${arg}". Try --help.`);
		const name = arg.slice(2).replace(/-([a-z])/gu, (_m, c) => c.toUpperCase());
		options[name] = true;
	}
	return options;
}

/** True when npm output says the publish needs a one-time password / second factor. */
export function isOtpError(output) {
	return /\bEOTP\b|one-time pass|\bOTP\b.*required|requires? (?:a )?(?:one-time|two-factor|2fa)|two-factor authentication|--otp/iu.test(
		String(output),
	);
}

/**
 * What an `npm whoami` result means: `ok` (logged in), `login` (nobody is logged in) or `error`
 * (the registry could not be reached, so logging in would not help).
 * @returns {'ok' | 'login' | 'error'}
 */
export function loginState(result) {
	if (result.status === 0 && String(result.stdout ?? '').trim()) return 'ok';
	const output = `${result.stdout ?? ''}${result.stderr ?? ''}`;
	if (/ENOTFOUND|ECONNREFUSED|ECONNRESET|ETIMEDOUT|EAI_AGAIN|network/iu.test(output))
		return 'error';
	return 'login';
}

/** Only the literal word `yes` (any case, surrounding spaces ignored) confirms. */
export const isConfirmed = (answer) =>
	String(answer ?? '')
		.trim()
		.toLowerCase() === 'yes';

/**
 * Rows of the publish plan: the packages the release plan marks `release` (optionally only
 * `keys`), each with whether that exact version is already on npm.
 * @param {{ order: string[], packages: Record<string, {release: boolean, npm: string, dir: string, version: string}> }} plan
 * @param {(name: string, version: string) => 'exists' | 'missing'} state
 */
export function buildPublishPlan(plan, state, keys = []) {
	return resolveTargets({ plan })
		.filter((target) => keys.length === 0 || keys.includes(target.key))
		.map((target) => ({ ...target, onNpm: state(target.npm, target.version) === 'exists' }));
}

/** Text table of the plan. */
export function formatPlan(rows) {
	if (rows.length === 0)
		return 'Nothing to publish: the release plan marks no package for release.';
	const width = Math.max(...rows.map((r) => `${r.npm}@${r.version}`.length));
	return rows
		.map(
			(r) =>
				`  ${`${r.npm}@${r.version}`.padEnd(width)}  ${r.onNpm ? 'already on npm (skipped)' : 'will be published'}`,
		)
		.join('\n');
}

const shell = process.platform === 'win32';
const sh = (command, args, options = {}) =>
	spawnSync(command, args, { cwd: ROOT, encoding: 'utf8', shell, ...options });

function step(title, command, args) {
	console.log(`\n> ${[command, ...args].join(' ')}  (${title})`);
	const result = sh(command, args, { stdio: 'inherit' });
	if (result.status !== 0) throw new Error(`${title} failed; nothing was published.`);
}

function whoami() {
	return sh('npm', ['whoami', '--registry', REGISTRY], { stdio: ['ignore', 'pipe', 'pipe'] });
}

async function ensureLogin({ dryRun }) {
	let result = whoami();
	const state = loginState(result);
	if (state === 'ok') {
		console.log(`Logged in to npm as ${result.stdout.trim()}.`);
		return;
	}
	if (state === 'error') throw new Error(`Could not reach ${REGISTRY}:\n${result.stderr}`);
	if (dryRun) {
		console.log('Not logged in to npm; a dry run does not need it, continuing.');
		return;
	}
	console.log(
		'You are not logged in to npm. Starting `npm login --auth-type=web` (it opens your browser).',
	);
	const login = sh('npm', ['login', '--auth-type=web', '--registry', REGISTRY], {
		stdio: 'inherit',
	});
	if (login.status !== 0) throw new Error('npm login did not complete; nothing was published.');
	result = whoami();
	if (loginState(result) !== 'ok') throw new Error('Still not logged in to npm after `npm login`.');
	console.log(`Logged in to npm as ${result.stdout.trim()}.`);
}

function ensureCleanMain({ allowDirty }) {
	const status = sh('git', ['status', '--porcelain'], { stdio: ['ignore', 'pipe', 'pipe'] });
	const branch = sh('git', ['branch', '--show-current'], { stdio: ['ignore', 'pipe', 'pipe'] });
	if (status.status !== 0 || branch.status !== 0) throw new Error('Not a git work tree.');
	const problems = [];
	if (status.stdout.trim()) problems.push('the work tree has uncommitted changes');
	if (branch.stdout.trim() !== 'main')
		problems.push(`the branch is "${branch.stdout.trim() || '(detached)'}", not main`);
	if (problems.length === 0) return;
	if (allowDirty) {
		console.warn(`Warning: ${problems.join(' and ')} (--allow-dirty).`);
		return;
	}
	throw new Error(
		`Refusing to publish: ${problems.join(' and ')}. Pass --allow-dirty to override.`,
	);
}

/**
 * Runs `npm publish`; returns status and output. In a terminal npm keeps stdin so its browser
 * (web) authentication can run: it prints a URL, opens the browser and waits for you to sign in.
 * `--auth-type=web` is always passed, so npm never falls back to asking for a one-time code.
 * Without a terminal stdin is closed and npm can only fail.
 */
function npmPublish(args, cwd) {
	return new Promise((resolve) => {
		const stdin = process.stdin.isTTY ? 'inherit' : 'ignore';
		const child = spawn('npm', [...args, '--auth-type=web'], {
			cwd,
			shell,
			stdio: [stdin, 'pipe', 'pipe'],
		});
		let output = '';
		const forward = (stream) => (chunk) => {
			output += chunk;
			stream.write(chunk);
		};
		child.stdout.on('data', forward(process.stdout));
		child.stderr.on('data', forward(process.stderr));
		child.on('close', (status) => resolve({ status, output }));
	});
}

async function confirm({ yes }) {
	if (yes) return true;
	if (!process.stdin.isTTY) throw new Error('No terminal to confirm in; pass --yes to publish.');
	const rl = createInterface({ input: process.stdin, output: process.stdout });
	try {
		return isConfirmed(await rl.question('\nType "yes" to publish these packages: '));
	} finally {
		rl.close();
	}
}

/** The tags release.yml expects for what was just published, so CI does not re-release it. */
export function tagCommands(rows) {
	const tags = rows.map((row) => `${row.npm}@${row.version}`);
	return tags.length === 0
		? []
		: [...tags.map((tag) => `git tag ${tag}`), `git push origin ${tags.join(' ')}`];
}

async function main() {
	const options = parseArgs(process.argv.slice(2));
	if (options.help) {
		console.log(
			'Usage: bun run publish:local -- [--dry-run] [--yes] [--allow-dirty] [--skip-build] [--package <key>]...\n' +
				'See docs/releasing.md, "Publishing from your machine".',
		);
		return;
	}
	await ensureLogin(options);
	ensureCleanMain(options);

	const plan = planRelease({
		root: ROOT,
		packages: PACKAGES,
		globalTriggers: GLOBAL_TRIGGERS,
		npm: npmVersion,
	});
	const rows = buildPublishPlan(plan, registryState, options.packages);
	// The same manifest checks CI runs: versions on disk match the plan, nothing local or private.
	for (const row of rows) verifyManifest(row);
	if (rows.every((row) => row.onNpm)) {
		console.log(formatPlan(rows));
		console.log('\nNothing new to publish.');
		return;
	}

	if (!options.skipBuild) step('build', 'bun', ['run', 'build:packages']);
	step('published-import check', 'bun', ['run', 'check:published']);
	step('pack smoke test', 'bun', ['run', 'pack:smoke']);

	const provenance = provenanceEnabled();
	console.log(`\nPublish plan (registry ${REGISTRY}, provenance ${provenance ? 'on' : 'off'}):`);
	console.log(formatPlan(rows));
	if (!options.dryRun && !(await confirm(options))) {
		console.log(
			'Not confirmed; nothing was published. If the prompt closed without letting you type, ' +
				'run it again with: bun run publish:local -- --yes --skip-build',
		);
		process.exitCode = 1;
		return;
	}

	const published = [];
	for (const row of rows) {
		if (row.onNpm) continue;
		const args = publishArgs({
			tag: distTag(row.npm, row.version),
			provenance,
			dryRun: options.dryRun,
		});
		console.log(`\n--- ${row.npm}@${row.version}: npm ${args.join(' ')} ---`);
		const result = await npmPublish(args, join(ROOT, row.dir));
		if (result.status === 0) {
			published.push(row);
			continue;
		}
		if (isOtpError(result.output)) {
			throw new Error(
				`npm still wants a one-time password to publish ${row.npm}, which publish:local never ` +
					'asks for. Run `npm login --auth-type=web` again (finish the sign-in in the browser ' +
					'before it times out), or publish with a granular access token that has "bypass 2FA" ' +
					'enabled (npmjs.com, Access Tokens), or use the CI trusted-publishing path (release.yml).',
			);
		}
		throw new Error(`Publishing ${row.npm}@${row.version} failed (exit ${result.status}).`);
	}
	if (options.dryRun) {
		console.log('\nDry run complete; nothing was published.');
		return;
	}
	console.log(
		'\nPublished. Tag the commit so the release workflow treats these versions as released:',
	);
	for (const command of tagCommands(published)) console.log(`  ${command}`);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
	main().catch((error) => {
		console.error(`\npublish:local: ${error.message}`);
		process.exit(1);
	});
}
