#!/usr/bin/env node
// Decides what CI has to run for a change, so a docs edit does not run the unit suite and a
// one-file change does not start six runners.
//
//   node scripts/ci-plan.mjs --base <sha> --event <push|pull_request|schedule|workflow_dispatch>
//
// Prints the plan as JSON and, with GITHUB_OUTPUT set, writes it as the `plan` output. The planner is
// deliberately conservative: anything it cannot place (a root manifest, a tsconfig, this file, the
// workflow) runs everything, and the nightly run is always a full one, so a missed dependency edge
// in `vitest --changed` is caught within a day instead of never.
import { execFileSync } from 'node:child_process';
import { appendFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const MAX_SHARDS = 6;
/** Test files one shard should take before another runner is worth its start-up cost. */
export const FILES_PER_SHARD = 60;
const ZERO_SHA = /^0+$/;

/** A copy of a viewer's standalone MCP server, installed on its own from the registry as a user would. */
const STANDALONE_MCP =
	'tmp=$(mktemp -d) && cp -r mcp/. "$tmp" && (cd "$tmp" && npm install --ignore-scripts --no-package-lock && npm test)';

/**
 * The viewers that live in this repository under `viewers/<name>`, and the commands that verify
 * each (run in order from the viewer's directory after the core and `ooxml-ui` are built, because
 * the viewers resolve `ooxml-ui` to the workspace copy through its `dist`). They replace the
 * CI each viewer had when it was a repository of its own.
 */
export const VIEWERS = {
	pptx: {
		name: 'pptx',
		dir: 'viewers/pptx',
		// More than one runner fits in the viewers job's ceiling (a ~20-minute unit suite and a
		// browser suite of hours of runner time), so pptx has its own jobs in ci.yml (`pptx-*`):
		// one build whose output the rest share, a unit job per package, the browser suite by
		// framework and shard, and the packaged-build smoke. `ownJobs` keeps it out of `viewers`;
		// `verify` is what its build job runs. Its `test:scripts` stays out: those scripts ran the
		// CI and releases of its own repository.
		ownJobs: true,
		verify: [
			'bun run build:packages',
			'bun run typecheck',
			'bun run e2e:contract',
			'bun run test:binding-packages',
		],
	},
	docx: {
		name: 'docx',
		dir: 'viewers/docx',
		verify: [
			'bun run typecheck',
			'bun run test',
			'bun run test:scripts',
			STANDALONE_MCP,
			'bun x playwright install --with-deps chromium',
			'bun run test:browser',
			'bun run build:packages',
			'bun run check:published',
			'bun run pack:smoke',
			'bun run check:shared',
		],
	},
	xlsx: {
		name: 'xlsx',
		dir: 'viewers/xlsx',
		verify: [
			'bun run typecheck',
			'bun run test',
			'bun run test:scripts',
			STANDALONE_MCP,
			'bun x playwright install --with-deps chromium',
			'bun run test:browser',
			'bun run build:packages',
			'bun run check:published',
			'bun run pack:smoke',
			'bun run check:shared',
		],
	},
	visio: {
		name: 'visio',
		dir: 'viewers/visio',
		// `check` runs the formatter, the core and converter checks, types, tests, the bindings, the
		// docs tests, both builds and the tarball tests.
		verify: [
			'bun run check',
			STANDALONE_MCP,
			'bun x playwright install --with-deps chromium',
			'bun run test:browser',
		],
	},
	teams: {
		name: 'teams',
		dir: 'viewers/teams',
		// The demos resolve the framework packages through their built dist, so build before types.
		verify: [
			'bun run build:packages',
			'bun run typecheck',
			'bun run test',
			'bun run test:scripts',
			'bun x playwright install --with-deps chromium',
			'bun run test:browser',
			'bun run check:published',
			'bun run pack:smoke',
		],
	},
};
const ALL_VIEWERS = Object.keys(VIEWERS);

/**
 * The viewer repositories that still consume this one, and how each is verified against a local
 * link (`{ name, repo, install, verify }`). None today: pptx-viewer, the last one, moved into
 * `viewers/pptx`. The `consumers` job stays for the next external consumer.
 */
export const CONSUMERS = {};
const ALL_CONSUMERS = Object.keys(CONSUMERS);

/** Files whose change can affect everything: dependencies, compiler and test configuration, CI. */
const EVERYTHING =
	/^(package\.json|bun\.lockb?|\.oxfmtrc\.json|\.github\/workflows\/ci\.yml|scripts\/ci-plan\.(mjs|test\.mjs)|src\/core\/(package\.json|tsconfig[^/]*\.json|vitest\.[^/]+|tsup[^/]*\.ts|tsdown[^/]*\.ts))$/;
/** Files that change nothing CI checks. */
const INERT = /^(docs\/|[^/]+\.md$|LICENSE|NOTICE|\.gitignore|\.gitattributes|\.claude\/|memory\/)/;
/** Top-level files of a viewer that change nothing CI checks (its docs are checked by its own tests). */
const VIEWER_INERT = /^viewers\/[^/]+\/([^/]+\.md|LICENSE|NOTICE|\.gitignore|\.gitattributes)$/;
/**
 * The viewer a path belongs to, or undefined: its own folder under `viewers/`, and the demos and
 * browser tests the viewers keep at the root (`demos/<viewer>/`, `e2e/<viewer>/`).
 */
const viewerOfFile = (file) => /^(?:viewers|demos|e2e)\/([^/]+)\//.exec(file)?.[1];

/** The area a format owns: only its own viewer can break. */
const AREA_CONSUMERS = { docx: [], xlsx: [], visio: [], teams: [], pptx: [] };
const AREA_VIEWERS = {
	docx: ['docx'],
	xlsx: ['xlsx'],
	visio: ['visio'],
	teams: ['teams'],
	pptx: ['pptx'],
};

/** Which external viewers a change under `src/core/<area>/` can break. Shared areas can break all of them. */
export function consumersOfArea(area) {
	return AREA_CONSUMERS[area] ?? ALL_CONSUMERS;
}

/** Which in-repo viewers a change under `src/core/<area>/` can break. Shared areas can break all of them. */
export function viewersOfArea(area) {
	return AREA_VIEWERS[area] ?? ALL_VIEWERS;
}

/** Shard indexes for `count` test files: none for zero, one per `FILES_PER_SHARD`, at most six. */
export function shardsFor(count, full) {
	if (full) return Array.from({ length: MAX_SHARDS }, (_, i) => i + 1);
	if (count <= 0) return [];
	const n = Math.min(MAX_SHARDS, Math.max(1, Math.ceil(count / FILES_PER_SHARD)));
	return Array.from({ length: n }, (_, i) => i + 1);
}

/** The pptx viewer's packages and the ones each depends on, for the unit legs a change reaches. */
const PPTX_DEPENDS_ON = {
	core: [],
	tools: ['core'],
	shared: ['core', 'tools'],
	locales: ['shared'],
	react: ['core', 'shared', 'locales'],
	react18: ['react'],
	vue: ['core', 'shared', 'locales'],
	angular: ['core', 'shared', 'locales'],
	vanilla: ['core', 'shared', 'locales'],
	svelte: ['core', 'shared', 'locales'],
	cli: ['core', 'react'],
};
const PPTX_LEGS = Object.keys(PPTX_DEPENDS_ON);
/** The five bindings, which are also the Playwright projects (one demo each). */
export const PPTX_BINDINGS = ['react', 'vue', 'angular', 'vanilla', 'svelte'];
/**
 * The Playwright project a cross-binding change runs. React is the parity reference: its parity
 * specs drive all five demos and diff them against it. The other four run for their own changes,
 * and all five on the nightly and on changes to dependencies or CI.
 */
const PPTX_REFERENCE_PROJECT = 'react';
/** Shards per Playwright project for a full run of the suite (pptx-viewer settled on ten). */
export const PPTX_E2E_SHARDS = 10;
/** Changed spec files one shard takes when only specs changed. */
const PPTX_SPECS_PER_SHARD = 5;

/** The legs that test `leg`'s code: itself and every leg depending on it, transitively. */
function pptxDependents(leg) {
	const out = new Set([leg]);
	let grew = true;
	while (grew) {
		grew = false;
		for (const [key, deps] of Object.entries(PPTX_DEPENDS_ON)) {
			if (!out.has(key) && deps.some((dep) => out.has(dep))) {
				out.add(key);
				grew = true;
			}
		}
	}
	return out;
}

const isTestFile = (path) => /(\.test\.[cm]?[jt]sx?|\/__tests__\/)/.test(path);

/**
 * What the pptx jobs run for a change. `reaches` is true when the change reaches the viewer at all
 * (its own files, an area of the core it reads, `ooxml-ui`, or everything). A change inside one
 * package runs that package's unit leg and its dependents' and, unless it only touched tests, the
 * browser suite for the binding it belongs to (the parity reference for shared packages). A change
 * to spec files alone runs just those specs. Configuration, fixtures and support code the whole
 * suite depends on run all of it.
 */
export function pptxPlan(files, { everything, reaches }) {
	const none = {
		run: false,
		tests: [],
		e2e: { projects: [], shards: [], total: 0, specs: '' },
		packaged: false,
	};
	if (!reaches) return none;
	const legs = new Set();
	const projects = new Set();
	const specs = new Set();
	let allLegs = everything;
	let allProjects = everything;
	let packaged = everything;
	for (const file of files) {
		if (file.startsWith('src/core/') || file.startsWith('src/ui/')) {
			// The engine or the shared elements: every package's tests, the parity reference's browser
			// run. Their own tests change nothing the viewer runs.
			if (isTestFile(file)) continue;
			allLegs = true;
			projects.add(PPTX_REFERENCE_PROJECT);
			packaged = true;
			continue;
		}
		let match;
		if ((match = /^viewers\/pptx\/packages\/([^/]+)\//.exec(file))) {
			const leg = match[1] === 'react-compat' ? 'react18' : match[1];
			if (!PPTX_DEPENDS_ON[leg]) continue;
			for (const dependent of pptxDependents(leg)) legs.add(dependent);
			if (isTestFile(file) || leg === 'cli') continue;
			projects.add(PPTX_BINDINGS.includes(leg) ? leg : PPTX_REFERENCE_PROJECT);
			packaged = true;
		} else if (/^viewers\/pptx\/(docs|\.github)\//.test(file)) {
			// The docs site is built by the Pages workflow; the old repository's workflows are inert.
		} else if ((match = /^demos\/pptx\/demo-([^/]+)\//.exec(file))) {
			// Each demo is its binding's browser surface; the locale and vue/vanilla tests import demos.
			if (PPTX_BINDINGS.includes(match[1])) projects.add(match[1]);
			for (const leg of ['locales', 'vue', 'vanilla']) legs.add(leg);
			packaged = true;
		} else if (file.startsWith('demos/pptx/')) {
			// Files every demo shares.
			allProjects = true;
			for (const leg of ['locales', 'vue', 'vanilla']) legs.add(leg);
			packaged = true;
		} else if (/^e2e\/pptx\/[^/]+\.spec\.ts$/.test(file)) {
			specs.add(file.slice('e2e/pptx/'.length));
		} else if (file.startsWith('e2e/pptx/fixtures/')) {
			// The fixtures are the demos' public dir and several packages' test decks.
			allProjects = true;
			for (const leg of ['shared', 'react', 'react18', 'vue', 'angular', 'vanilla', 'svelte'])
				legs.add(leg);
		} else if (file.startsWith('e2e/pptx/') || file.startsWith('viewers/pptx/')) {
			// Support code, configs, scripts and manifests the whole viewer depends on.
			allLegs = true;
			allProjects = true;
			packaged = true;
		}
	}
	const tests = allLegs ? PPTX_LEGS : PPTX_LEGS.filter((leg) => legs.has(leg));
	let projectList = allProjects ? PPTX_BINDINGS : PPTX_BINDINGS.filter((p) => projects.has(p));
	let total = projectList.length > 0 ? PPTX_E2E_SHARDS : 0;
	let specFilter = '';
	if (projectList.length === 0 && specs.size > 0) {
		// Only specs changed: run just those, in every binding, on as few shards as they need.
		projectList = PPTX_BINDINGS;
		total = Math.min(PPTX_E2E_SHARDS, Math.ceil(specs.size / PPTX_SPECS_PER_SHARD));
		specFilter = [...specs].sort().join(' ');
	}
	const run = tests.length > 0 || projectList.length > 0 || packaged;
	return {
		run,
		tests,
		e2e: {
			projects: projectList,
			shards: Array.from({ length: total }, (_, i) => i + 1),
			total,
			specs: specFilter,
		},
		packaged,
	};
}

/**
 * The plan for a list of changed files. `full` runs everything (nightly, manual, or an unknown base).
 * `testFiles` is the number of test files `vitest --changed` selected, or undefined when it has not
 * been counted (the planner then reports a placeholder and the workflow counts).
 */
export function plan(changed, { full = false, testFiles } = {}) {
	const files = changed.filter(Boolean);
	const everything = full || files.some((file) => EVERYTHING.test(file));
	const live = files.filter((file) => !INERT.test(file) && !VIEWER_INERT.test(file));
	// src/core is the library, one folder per area; src/ui is ooxml-ui, a package of its own.
	const src = live.filter((file) => file.startsWith('src/core/'));
	const areas = new Set(src.map((file) => file.split('/')[2]));
	const ui = live.some((file) => file.startsWith('src/ui/'));
	const scripts = live.some((file) => /^(scripts|\.github|site)\//.test(file));
	const mcp = live.some(
		(file) => file.startsWith('mcp/') || file.startsWith('src/core/automation/'),
	);
	const core = src.length > 0;

	const consumers = new Set();
	if (everything || ui) for (const key of ALL_CONSUMERS) consumers.add(key);
	for (const area of areas) for (const key of consumersOfArea(area)) consumers.add(key);

	// A viewer is checked when its own files change, when the shared UI or an area it reads changes,
	// and on everything. Its scripts read the shared package table, so that file counts too (visio).
	const viewers = new Set();
	if (everything || ui) for (const key of ALL_VIEWERS) viewers.add(key);
	for (const area of areas) for (const key of viewersOfArea(area)) viewers.add(key);
	for (const file of live) {
		const owner = viewerOfFile(file);
		if (owner && VIEWERS[owner]) viewers.add(owner);
		if (file === 'scripts/viewer-packages.mjs') viewers.add('visio');
	}
	const viewerList = ALL_VIEWERS.filter((key) => viewers.has(key)).map((key) => VIEWERS[key]);

	const testsNeeded = everything || core;
	return {
		full: everything,
		typecheck: {
			strict: everything || src.some((file) => !file.startsWith('src/core/pptx/')),
			pptx: everything || core,
			ui: everything || ui || core,
		},
		test: {
			run: testsNeeded,
			mode: everything ? 'all' : 'changed',
			shards: testsNeeded ? shardsFor(testFiles ?? MAX_SHARDS, everything) : [],
		},
		// The viewers resolve ooxml-ui to the workspace copy through its `dist`, so any viewer run needs
		// the core and the UI built.
		ui: everything || ui || core || viewerList.length > 0,
		build: everything || core || ui || viewerList.length > 0,
		scripts: everything || scripts,
		mcp: everything || mcp,
		viewers: viewerList
			.filter((viewer) => !viewer.ownJobs)
			.map(({ name, dir, verify }) => ({ name, dir, verify: verify.join('\n') })),
		// What the pptx jobs in ci.yml run (see VIEWERS.pptx and pptxPlan).
		pptx: pptxPlan(live, {
			everything,
			reaches: viewerList.some((viewer) => viewer.name === 'pptx'),
		}),
		consumers: ALL_CONSUMERS.filter((key) => consumers.has(key)).map((key) => CONSUMERS[key]),
	};
}

const git = (...args) => execFileSync('git', args, { encoding: 'utf8' }).trim();

/** The files changed since `base`, or undefined when the base is unusable (new branch, shallow). */
export function changedSince(base) {
	if (!base || ZERO_SHA.test(base)) return undefined;
	try {
		return git('diff', '--name-only', `${base}...HEAD`).split('\n');
	} catch {
		try {
			return git('diff', '--name-only', base, 'HEAD').split('\n');
		} catch {
			return undefined;
		}
	}
}

function countTestFiles(base) {
	const out = execFileSync('bunx', ['vitest', 'list', '--changed', base, '--filesOnly'], {
		cwd: 'src/core',
		encoding: 'utf8',
		shell: process.platform === 'win32',
	});
	return out.split('\n').filter((line) => /\.(test|spec)\.[cm]?[tj]sx?$/.test(line.trim())).length;
}

function main() {
	const args = process.argv.slice(2);
	const value = (flag) => args[args.indexOf(flag) + 1];
	const event = value('--event') ?? 'push';
	const base = value('--base');
	const changed =
		event === 'schedule' || event === 'workflow_dispatch' ? undefined : changedSince(base);
	const full = changed === undefined;
	let result = plan(changed ?? [], { full });
	if (!result.full && result.test.run) {
		const count = countTestFiles(base);
		result = { ...result, test: { ...result.test, shards: shardsFor(count, false), files: count } };
	}
	result.base = full ? '' : base;
	console.log(JSON.stringify(result, null, 2));
	if (process.env.GITHUB_OUTPUT)
		appendFileSync(process.env.GITHUB_OUTPUT, `plan=${JSON.stringify(result)}\n`);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) main();
