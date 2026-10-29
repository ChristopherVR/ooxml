#!/usr/bin/env node
/**
 * verify.mjs: run locally what CI would run for the commits not yet on the base branch.
 *
 *   node scripts/verify.mjs [--base origin/main] [--browser] [--full] [--allow-dirty]
 *                           [--jobs N | --serial] [--rebuild]
 *
 * The checks come from the CI planner (scripts/ci-plan.mjs), so a change runs the same lint,
 * typechecks, tests and viewer checks here as in CI. By default the slow groups are skipped:
 * browser suites, package packing and the standalone MCP install (`--browser` adds the browser
 * suites, `--full` adds everything and checks the whole repository, as the nightly run does).
 * Every step runs with CI=1, so Vitest fails on a missing snapshot instead of writing it.
 *
 * Independent steps run in parallel (scripts/verify-runner.mjs): `--jobs N` sets how many, the
 * default is a quarter of the cores because vitest and tsc are multi-threaded themselves, and
 * `--serial` runs them one by one in plan order. Each step's output is buffered and printed whole
 * when it finishes. After a failure nothing new starts, running steps finish, and the failed
 * step's output is printed last.
 *
 * The core and ooxml-ui builds are skipped when their tracked inputs are unchanged since the last
 * successful build here (a stamp in `dist`, see `hashInputs`); `--rebuild` forces them. Untracked
 * files are not part of that hash, which is sound because verify refuses a dirty tree.
 *
 * It refuses to run on uncommitted changes to tracked files: the result would describe the working
 * tree, not the commits being pushed (and in a shared checkout, possibly someone else's edits).
 * A local pre-push hook can run it (see AGENTS.md); `git push --no-verify` skips it once.
 */
import { spawn, execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { availableParallelism } from 'node:os';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PLAYWRIGHT_INSTALL, changedSince, plan } from './ci-plan.mjs';
import { argumentBatches, FORMATTER_CHECK } from './verify-arguments.mjs';
import { browserPerformance, pptxBrowserCommand } from './browser-performance.mjs';
import { verificationShell } from './verify-shell.mjs';
import {
	clearStamp,
	defaultJobs,
	hashInputs,
	runSteps,
	stampMatches,
	writeStamp,
} from './verify-runner.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const argv = process.argv.slice(2);
const flag = (name) => argv.includes(name);
const value = (name, fallback) => (argv.includes(name) ? argv[argv.indexOf(name) + 1] : fallback);
const full = flag('--full');
const browser = full || flag('--browser');
const base = value('--base', 'origin/main');
const git = (...args) => execFileSync('git', args, { cwd: ROOT, encoding: 'utf8' }).trim();

/** Commands of a viewer's verify list that only `--browser` or `--full` runs. */
const BROWSER = /playwright|test:browser|e2e/;
const SLOW =
	/mktemp|pack:smoke|build:packages|test:packages|test-packages\.mjs|check:published|test:package\b|test:worker/;
const FORMATTED = /\.(m?[jt]sx?|css|json|md|vue|svelte|html|ya?ml)$/;

function fail(message) {
	console.error(`\nverify: ${message}`);
	process.exit(1);
}

if (!flag('--allow-dirty') && git('status', '--porcelain', '--untracked-files=no'))
	fail(
		'tracked files have uncommitted changes, so the checks would not describe the commits being\n' +
			'pushed. Commit or stash them (or verify from a clean worktree), or pass --allow-dirty.',
	);

// CI installs with the Bun version of .github/actions/setup; another version can write a
// different lockfile and resolve differently.
const action = readFileSync(join(ROOT, '.github/actions/setup/action.yml'), 'utf8');
const ciBun = /bun-version:[\s\S]*?default:\s*'([^']+)'/.exec(action)?.[1];
let localBun = '';
try {
	// Through the shell: on Windows `bun` is a shim Node cannot spawn by name.
	localBun = execFileSync('bun --version', { encoding: 'utf8', shell: true }).trim();
} catch {
	console.warn('verify: warning: could not run `bun --version`.');
}
if (ciBun && localBun && ciBun !== localBun)
	console.warn(`verify: warning: local Bun ${localBun} differs from CI's ${ciBun} (bun upgrade).`);

const changed = full ? [] : changedSince(base);
if (!full && !changed) fail(`cannot diff against ${base}; fetch it or pass --base <ref>.`);
const files = (changed ?? []).filter(Boolean);
if (!full && !files.length) {
	console.log(`verify: nothing differs from ${base}.`);
	process.exit(0);
}
const p = plan(files, { full });

// ---- Steps. Each has an id and the ids it needs (see verify-runner.mjs). ----------------------
const steps = [];
const add = (id, cmd, cwd = '.', needs = [], extra = {}) => {
	steps.push({ id, cmd, cwd, needs, ...extra });
	return id;
};

/** A build skipped while its inputs hash equals the stamp in its `dist`. */
function buildStep(id, cmd, cwd, dir, { extra = [], needs = [] } = {}) {
	const dist = join(ROOT, dir, 'dist');
	const hash = hashInputs({
		root: ROOT,
		dirs: [dir],
		files: ['bun.lock', 'package.json'],
		extra,
		lsFiles: (path) => git('ls-files', '-s', '--', path),
	});
	add(id, cmd, cwd, needs, {
		hash,
		skipIf: () => !flag('--rebuild') && stampMatches(dist, hash) && 'inputs unchanged',
		before: () => clearStamp(dist),
		onSuccess: () => writeStamp(dist, hash),
	});
	return hash;
}

add('lint', 'bun run lint');

// Formatting of the changed files, with each viewer's own oxfmt configuration.
const present = files.filter((file) => FORMATTED.test(file) && existsSync(join(ROOT, file)));
const byViewer = new Map();
for (const file of present) {
	const dir = /^viewers\/[^/]+/.exec(file)?.[0] ?? '.';
	byViewer.set(dir, [...(byViewer.get(dir) ?? []), dir === '.' ? file : relative(dir, file)]);
}
let fmtCount = 0;
for (const [dir, list] of byViewer)
	for (const command of argumentBatches(FORMATTER_CHECK, list))
		add(`fmt:${fmtCount++}`, command, dir);

// The builds the viewers and the ooxml-ui steps consume. Typechecks and unit tests of the core read
// the source (vitest aliases the core to it), so they do not wait.
const BUILT = [];
if (p.build) {
	// After lint: oxlint's import plugin resolves `ooxml-core/*` into `src/core/dist`, and a build
	// that replaces `dist` while lint reads it fails lint with "Failed to open file".
	const coreHash = buildStep('build:core', 'bun run build', '.', 'src/core', { needs: ['lint'] });
	buildStep('build:ui', 'bun run --cwd src/ui build', '.', 'src/ui', {
		extra: [coreHash],
		needs: ['build:core'],
	});
	BUILT.push('build:core', 'build:ui');
}
if (p.typecheck.strict) add('typecheck:strict', 'bun run typecheck:strict');
if (p.typecheck.pptx) add('typecheck:pptx', 'bun run typecheck:pptx');
// ooxml-ui resolves the built core (package exports), unlike its vitest suite.
if (p.typecheck.ui) add('typecheck:ui', 'bun run --cwd src/ui typecheck', '.', ['build:core']);
const changedTests = full ? '' : ` --changed=${base} --passWithNoTests`;
// The core suite follows the plan (a manifest or config change runs all of it, as in CI). The
// ooxml-ui suite takes about a quarter of an hour, so only `--full` runs it whole; CI always does.
if (p.test.run) add('test:core', `bun run test${p.test.mode === 'all' ? '' : changedTests}`);
if (p.checks.ui)
	// Measured: without `src/core/dist` a pptx test in the suite fails to resolve `ooxml-core/pptx`.
	add('test:ui', `bunx vitest run${changedTests}`, 'src/ui', ['build:core']);
if (p.scripts) add('test:scripts', 'bun run test:scripts');
if (p.mcp) add('test:mcp', 'bun run test:mcp');
for (const viewer of p.viewers) {
	// Within a viewer the commands keep their order (they have some, for example a build before a
	// typecheck); different viewers run side by side once the builds are there.
	let previous = BUILT;
	for (const [index, cmd] of viewer.verify.split('\n').entries())
		if ((browser || !BROWSER.test(cmd)) && (full || !SLOW.test(cmd)))
			previous = [add(`${viewer.name ?? viewer.dir}:${index}`, cmd, viewer.dir, previous)];
}
// The browser suites are separate, sharded CI entries; locally one run covers each whole suite,
// after the viewer's own checks. Suites run side by side, and docx and xlsx share a default
// preview port, so each gets its own (`PLAYWRIGHT_PORT`, read by every viewer's config).
if (browser)
	for (const [suite, entry] of (p.viewerBrowser ?? [])
		.filter((item) => item.shard === 1)
		.entries()) {
		const name = `${entry.name}:browser`;
		const port = `export PLAYWRIGHT_PORT=${4300 + suite * 20}; `;
		let previous = [add(`${name}:install`, PLAYWRIGHT_INSTALL, entry.dir, BUILT)];
		for (const [index, cmd] of entry.local.split('\n').entries())
			previous = [add(`${name}:${index}`, port + cmd, entry.dir, previous, { exclusive: true })];
	}
if (p.pptx.run) {
	// `packaged` is the planner's "this change alters what the viewer's packages build". Without it
	// (a change to tests or specs only) the typecheck trio, which reads the built packages, is not
	// needed, and neither is `build:packages`. Unknown means yes.
	const dist = p.pptx.packaged !== false || (browser && p.pptx.e2e?.projects?.length > 0);
	// Measured without the packages' dist: the vue, react, core and tools legs pass with only the
	// core and ui builds; the locales and cli legs import the built sibling packages.
	const legs = p.pptx.tests;
	const packages = dist || legs.some((leg) => leg === 'locales' || leg === 'cli');
	let last = BUILT;
	if (packages) {
		const cwd = 'viewers/pptx';
		last = [add('pptx:build', 'bun run build:packages', cwd, BUILT)];
		if (dist)
			for (const [name, cmd] of [
				['typecheck', 'bun run typecheck'],
				['viewer-props', 'bun run check:viewer-props'],
				['tree-shaking', 'bun run check:tree-shaking'],
			])
				add(`pptx:${name}`, cmd, cwd, last);
	}
	for (const leg of legs) {
		const needs = leg === 'locales' || leg === 'cli' ? [...BUILT, 'pptx:build'] : BUILT;
		if (leg === 'react18')
			add('pptx:test:react18', 'bun run test:react18', 'viewers/pptx/packages/react', needs);
		else add(`pptx:test:${leg}`, 'bun run test', `viewers/pptx/packages/${leg}`, needs);
	}
	if (browser && p.pptx.e2e?.projects?.length) {
		const demos = add('pptx:demos', 'bun e2e-build.ts', 'viewers/pptx', ['pptx:build', ...BUILT], {
			env: { PPTX_E2E_PROJECTS: p.pptx.e2e.projects.join(',') },
		});
		add('pptx:e2e', pptxBrowserCommand(p.pptx.e2e), 'viewers/pptx', [demos], {
			exclusive: true,
			env: { PPTX_E2E_SERVE: 'preview' },
		});
	}
}

// ---- Running. -------------------------------------------------------------------------------
// Commands are written for bash (the planner's commands use bash syntax, as in CI).
const shell = verificationShell();
// Git hands a hook the repository it runs for (GIT_DIR, GIT_INDEX_FILE and so on). The steps must
// not inherit it: the release planner's tests build throwaway repositories with `git init`,
// `commit`, `tag` and `config`, which would otherwise land in this repository and its shared
// config when verify runs as the pre-push hook. Each step finds the repository from its directory.
const REPOSITORY_ENV =
	/^GIT_(?:DIR|WORK_TREE|INDEX_FILE|COMMON_DIR|OBJECT_DIRECTORY|ALTERNATE_OBJECT_DIRECTORIES|PREFIX)$/;
const env = Object.fromEntries(
	Object.entries({
		...process.env,
		CI: '1',
		OOXML_LOCAL_VERIFY: '1',
		OOXML_E2E_WORKERS: String(
			browserPerformance({ ...process.env, OOXML_LOCAL_VERIFY: '1' }).workers,
		),
		PPTX_E2E_PORT_OFFSET: process.env.PPTX_E2E_PORT_OFFSET ?? '200',
	}).filter(([name]) => !REPOSITORY_ENV.test(name)),
);

const serial = flag('--serial');
const jobs = serial ? 1 : Number(value('--jobs', defaultJobs(availableParallelism())));
if (!Number.isInteger(jobs) || jobs < 1) fail('--jobs takes a positive integer.');
if (serial) {
	// Plan order, one at a time: the order the steps were added in already respects `needs`.
	for (const step of steps) step.needs = [];
}

const label = (step, index) =>
	`[${index + 1}/${steps.length}] ${step.cwd === '.' ? '' : `${step.cwd}: `}${step.cmd}`;
const labels = new Map(steps.map((step, index) => [step, label(step, index)]));

function exec(step) {
	step.before?.();
	return new Promise((resolveStep) => {
		const child = spawn(shell, ['-e', '-o', 'pipefail', '-s'], {
			cwd: join(ROOT, step.cwd),
			stdio: ['pipe', 'pipe', 'pipe'],
			env: { ...env, ...step.env },
		});
		const chunks = [];
		child.stdout.on('data', (chunk) => chunks.push(chunk));
		child.stderr.on('data', (chunk) => chunks.push(chunk));
		child.on('error', (error) => resolveStep({ code: 1, output: String(error) }));
		child.on('close', (code) =>
			resolveStep({ code: code ?? 1, output: Buffer.concat(chunks).toString() }),
		);
		child.stdin.end(step.cmd + '\n');
	});
}

const seconds = (ms) => `${(ms / 1000).toFixed(1)}s`;
const tty = process.stdout.isTTY && !flag('--no-status');
const startedAt = new Map();
let statusShown = false;
const clearStatus = () => {
	if (statusShown) process.stdout.write('\r\x1b[2K');
	statusShown = false;
};
const showStatus = () => {
	if (!tty) return;
	const now = Date.now();
	const line = [...startedAt]
		.map(([step, at]) => `${step.id} ${Math.round((now - at) / 1000)}s`)
		.join(' | ');
	process.stdout.write(
		`\r\x1b[2Kverify: running ${line}`.slice(0, (process.stdout.columns ?? 120) - 1),
	);
	statusShown = true;
};
const ticker = tty ? setInterval(showStatus, 1000) : undefined;

console.log(
	`verify: ${steps.length} steps (${jobs} at a time) for ${full ? 'the whole repository' : `${files.length} files changed since ${base}`}`,
);
const started = Date.now();
const { results, failed } = await runSteps(steps, {
	jobs,
	exec,
	onStart: (step) => {
		startedAt.set(step, Date.now());
		if (serial) console.log(`\n${labels.get(step)}`);
		showStatus();
	},
	onDone: (result) => {
		startedAt.delete(result.step);
		clearStatus();
		if (result.status === 'blocked') return showStatus();
		// A failed step's output is held back and printed last, below the summary.
		if (result.status === 'failed') return showStatus();
		if (!serial) console.log(`\n${labels.get(result.step)}`);
		if (result.output)
			process.stdout.write(result.output.endsWith('\n') ? result.output : `${result.output}\n`);
		if (result.status === 'ok') console.log(`ok in ${seconds(result.ms)}`);
		else console.log(`skipped (${result.reason})`);
		showStatus();
	},
});
if (ticker) clearInterval(ticker);
clearStatus();

const summary = [...results]
	.filter((result) => result.status !== 'blocked')
	.sort((a, b) => b.ms - a.ms)
	.map(
		(result) =>
			`  ${seconds(result.ms).padStart(8)}  ${result.status === 'ok' ? '' : `${result.status} `}${result.step.id}`,
	);
console.log(`\nverify: timings\n${summary.join('\n')}`);

if (failed) {
	for (const result of results.filter((entry) => entry.status === 'failed')) {
		console.error(`\n${labels.get(result.step)}\n${result.output}`);
	}
	const blocked = results.filter((result) => result.status === 'blocked').length;
	const first = results.find((result) => result.status === 'failed');
	fail(
		`failed: ${labels.get(first.step)}${blocked ? ` (${blocked} steps not run)` : ''}\n` +
			'Fix it, or push anyway with --no-verify.',
	);
}
console.log(
	`\nverify: all ${steps.length} steps passed in ${Math.round((Date.now() - started) / 1000)}s.`,
);
