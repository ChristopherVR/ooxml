#!/usr/bin/env node
/**
 * build-released.mjs: build (and optionally smoke-test) the workspace packages a release publishes.
 *
 * The root package (core) is built by `bun run build` / `bun run test:package` in the workflow;
 * this script covers every other package of release-plan.json (today `packages/ui`) that is
 * released in the plan, by running its own `build` (and, with `--smoke`, `test:package`) script
 * when it defines one. It runs after core is built because a workspace package resolves core
 * through its `dist`.
 *
 *   node scripts/build-released.mjs [--smoke] [--all] [--plan release-plan.json]
 *
 * `--all` ignores the `release` flag (build every workspace package present in the plan).
 */

import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const argv = process.argv.slice(2);
const planPath = argv.includes('--plan') ? argv[argv.indexOf('--plan') + 1] : 'release-plan.json';
const plan = JSON.parse(readFileSync(join(ROOT, planPath), 'utf8'));

for (const key of plan.order) {
	const pkg = plan.packages[key];
	if (pkg.dir === '.' || !(pkg.release || argv.includes('--all'))) continue;
	const scripts = JSON.parse(readFileSync(join(ROOT, pkg.manifest), 'utf8')).scripts ?? {};
	for (const script of argv.includes('--smoke') ? ['build', 'test:package'] : ['build']) {
		if (!scripts[script]) continue;
		console.log(`--- ${pkg.npm}: bun run ${script} ---`);
		const result = spawnSync('bun', ['run', script], {
			cwd: join(ROOT, pkg.dir),
			stdio: 'inherit',
			shell: process.platform === 'win32',
		});
		if (result.status !== 0) throw new Error(`${pkg.npm}: bun run ${script} failed.`);
	}
}
