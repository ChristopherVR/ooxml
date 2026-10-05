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

/** The viewer repositories that consume this one, and how each is verified against a local link. */
export const CONSUMERS = {
	docx: {
		name: 'docx-viewer',
		repo: 'ChristopherVR/docx-viewer',
		install: 'bun install',
		verify: 'bun run typecheck && bun run test',
	},
	xlsx: {
		name: 'xlsx-viewer',
		repo: 'ChristopherVR/xlsx-viewer',
		install: 'bun install',
		verify: 'bun run typecheck && bun run test',
	},
	visio: {
		name: 'visio-viewer',
		repo: 'ChristopherVR/visio-viewer',
		install: 'npm install',
		verify: 'npm run typecheck && npm test',
	},
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
	teams: {
		name: 'teams-viewer',
		repo: 'ChristopherVR/teams-viewer',
		install: 'bun install',
		// The demos resolve the framework packages through their built dist.
		verify: 'bun run build:packages && bun run typecheck && bunx vitest run',
	},
};
const ALL_CONSUMERS = Object.keys(CONSUMERS);

/** Files whose change can affect everything: dependencies, compiler and test configuration, CI. */
const EVERYTHING =
	/^(package\.json|bun\.lockb?|tsconfig[^/]*\.json|vitest\.[^/]+|\.oxfmtrc\.json|\.github\/workflows\/ci\.yml|scripts\/ci-plan\.(mjs|test\.mjs))$/;
/** Files that change nothing CI checks. */
const INERT = /^(docs\/|[^/]+\.md$|LICENSE|NOTICE|\.gitignore|\.gitattributes|\.claude\/|memory\/)/;
const AREA_CONSUMERS = {
	docx: ['docx'],
	xlsx: ['xlsx'],
	pptx: ['pptx'],
	visio: ['visio'],
	teams: ['teams'],
};

/** Which viewers a change under `src/<area>/` can break. Shared areas can break all of them. */
export function consumersOfArea(area) {
	return AREA_CONSUMERS[area] ?? ALL_CONSUMERS;
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
	const live = files.filter((file) => !INERT.test(file));
	const src = live.filter((file) => file.startsWith('src/'));
	const areas = new Set(src.map((file) => file.split('/')[1]));
	const ui = live.some((file) => file.startsWith('packages/ui/'));
	const scripts = live.some((file) => /^(scripts|\.github|site)\//.test(file));
	const mcp = live.some((file) => file.startsWith('mcp/') || file.startsWith('src/automation/'));
	const core = src.length > 0;

	const consumers = new Set();
	if (everything || ui) for (const key of ALL_CONSUMERS) consumers.add(key);
	for (const area of areas) for (const key of consumersOfArea(area)) consumers.add(key);

	const testsNeeded = everything || core;
	return {
		full: everything,
		typecheck: {
			strict: everything || src.some((file) => !file.startsWith('src/pptx/')),
			pptx: everything || core,
			ui: everything || ui || core,
		},
		test: {
			run: testsNeeded,
			mode: everything ? 'all' : 'changed',
			shards: testsNeeded ? shardsFor(testFiles ?? MAX_SHARDS, everything) : [],
		},
		ui: everything || ui || core,
		build: everything || core || ui,
		scripts: everything || scripts,
		mcp: everything || mcp,
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
