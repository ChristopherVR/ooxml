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
			'bun run check:published',
			'bun run pack:smoke',
		],
	},
};
const ALL_VIEWERS = Object.keys(VIEWERS);

/** The viewer repositories that still consume this one, and how each is verified against a local link. */
export const CONSUMERS = {
	pptx: {
		name: 'pptx-viewer',
		repo: 'ChristopherVR/pptx-viewer',
		install: 'bun install',
		// The shared package holds the pptx-ui-* subclasses of the ooxml-ui elements: the contract
		// that broke first every time an element changed. The core package is built first because shared
		// imports its dist. The full suite is the viewer's own CI.
		verify:
			'cd packages/core && bun run build && cd ../shared && bunx vitest run src/web-components',
	},
};
const ALL_CONSUMERS = Object.keys(CONSUMERS);

/** Files whose change can affect everything: dependencies, compiler and test configuration, CI. */
const EVERYTHING =
	/^(package\.json|bun\.lockb?|tsconfig[^/]*\.json|vitest\.[^/]+|\.oxfmtrc\.json|\.github\/workflows\/ci\.yml|scripts\/ci-plan\.(mjs|test\.mjs))$/;
/** Files that change nothing CI checks. */
const INERT = /^(docs\/|[^/]+\.md$|LICENSE|NOTICE|\.gitignore|\.gitattributes|\.claude\/|memory\/)/;
/** Top-level files of a viewer that change nothing CI checks (its docs are checked by its own tests). */
const VIEWER_INERT = /^viewers\/[^/]+\/([^/]+\.md|LICENSE|NOTICE|\.gitignore|\.gitattributes)$/;
/** The viewer a path under `viewers/` belongs to, or undefined. */
const viewerOfFile = (file) => /^viewers\/([^/]+)\//.exec(file)?.[1];

/** The area a format owns: only its own viewer (and, for pptx, the external consumer) can break. */
const AREA_CONSUMERS = { docx: [], xlsx: [], visio: [], teams: [], pptx: ['pptx'] };
const AREA_VIEWERS = {
	docx: ['docx'],
	xlsx: ['xlsx'],
	visio: ['visio'],
	teams: ['teams'],
	pptx: [],
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
		viewers: viewerList.map(({ name, dir, verify }) => ({ name, dir, verify: verify.join('\n') })),
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
