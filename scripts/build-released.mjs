#!/usr/bin/env node
/**
 * build-released.mjs: build (and optionally smoke-test) the workspace packages a release publishes.
 *
 * The root package (core) is built by `bun run build` / `bun run test:package` in the workflow;
 * this script covers every other package of release-plan.json that is released in the plan:
 *
 *   - a package under `viewers/<name>` is built by its viewer, once per viewer: the viewer's
 *     `build:packages` script produces every package of that viewer together (the framework
 *     bindings inline the viewer's private packages). With `--smoke` the viewer's
 *     `check:published` script (a static scan of what the tarballs import) runs afterwards. The
 *     viewers' `pack:smoke` / `test:packages` are not run here: they install the packed tarballs
 *     from the registry, and a dependency released in the same run is not on npm yet.
 *   - any other package (today `packages/ui`) runs its own `build` (and, with `--smoke`,
 *     `test:package`) script when it defines one. It runs after core is built because a workspace
 *     package resolves core through its `dist`.
 *
 *   node scripts/build-released.mjs [--smoke] [--all] [--plan release-plan.json]
 *
 * `--all` ignores the `release` flag (build every package present in the plan).
 */

import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { viewerOf } from './viewer-packages.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const argv = process.argv.slice(2);
const smoke = argv.includes('--smoke');
const planPath = argv.includes('--plan') ? argv[argv.indexOf('--plan') + 1] : 'release-plan.json';
const plan = JSON.parse(readFileSync(join(ROOT, planPath), 'utf8'));

const scriptsOf = (manifest) =>
	JSON.parse(readFileSync(join(ROOT, manifest), 'utf8')).scripts ?? {};

function run(label, cwd, script) {
	console.log(`--- ${label}: bun run ${script} ---`);
	const result = spawnSync('bun', ['run', script], {
		cwd: join(ROOT, cwd),
		stdio: 'inherit',
		shell: process.platform === 'win32',
	});
	if (result.status !== 0) throw new Error(`${label}: bun run ${script} failed.`);
}

const builtViewers = new Set();
for (const key of plan.order) {
	const pkg = plan.packages[key];
	if (pkg.dir === '.' || !(pkg.release || argv.includes('--all'))) continue;
	const viewer = viewerOf(pkg.dir);
	if (viewer) {
		if (builtViewers.has(viewer)) continue;
		builtViewers.add(viewer);
		const scripts = scriptsOf(`${viewer}/package.json`);
		if (scripts['build:packages']) run(viewer, viewer, 'build:packages');
		if (smoke && scripts['check:published']) run(viewer, viewer, 'check:published');
		continue;
	}
	const scripts = scriptsOf(pkg.manifest);
	for (const script of smoke ? ['build', 'test:package'] : ['build']) {
		if (scripts[script]) run(pkg.npm, pkg.dir, script);
	}
}
