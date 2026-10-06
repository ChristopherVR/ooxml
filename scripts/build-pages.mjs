#!/usr/bin/env node
/**
 * build-pages.mjs: build the whole GitHub Pages site into `pages-dist/`.
 *
 * One site serves https://christophervr.github.io/ooxml/: the Office-suite launcher (`site/`, static)
 * at the root, and each imported viewer's documentation and framework demos under `/<viewer>/`
 * (pptx, docx, xlsx, visio, teams). Every viewer builds itself with its own `scripts/build-pages.mjs`,
 * whose VitePress `base` and demo routes already live under `/ooxml/<viewer>/`; this script only
 * runs them and assembles the output.
 *
 *   node scripts/build-pages.mjs [--only docx,teams] [--out pages-dist]
 *
 * Needs `bun install`, the core built (`bun run build`) and `ooxml-ui` built
 * (`bun run --cwd src/ui build`), because the demos resolve both through their `dist`.
 */

import { spawnSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, rmSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { bundleSuite } from './bundle-suite.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const argv = process.argv.slice(2);
const value = (flag) => (argv.includes(flag) ? argv[argv.indexOf(flag) + 1] : undefined);

/** The viewers on the site, and what each needs before its own page build. */
export const VIEWERS = {
	pptx: { prepare: ['build:packages'] },
	docx: { prepare: ['build:packages'] },
	xlsx: { prepare: ['build:packages'] },
	visio: { prepare: [] },
	teams: { prepare: [] },
};

const only = value('--only')?.split(',');
const selected = Object.keys(VIEWERS).filter((name) => !only || only.includes(name));
const out = resolve(ROOT, value('--out') ?? 'pages-dist');

function run(command, args, cwd) {
	console.log(`--- ${cwd.slice(ROOT.length + 1) || '.'}: ${command} ${args.join(' ')} ---`);
	const result = spawnSync(command, args, {
		cwd,
		stdio: 'inherit',
		shell: process.platform === 'win32' && command === 'bun',
	});
	if (result.error) throw result.error;
	if (result.status !== 0) throw new Error(`${command} ${args.join(' ')} failed in ${cwd}`);
}

rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });

// The launcher goes first: it owns the site root, and no viewer directory name collides with it.
cpSync(join(ROOT, 'site'), out, { recursive: true });
await bundleSuite(join(out, 'suite.js'));

for (const name of selected) {
	const dir = join(ROOT, 'viewers', name);
	for (const script of VIEWERS[name].prepare) run('bun', ['run', script], dir);
	run(process.execPath, ['scripts/build-pages.mjs'], dir);
	const built = join(dir, 'docs', '.vitepress', 'dist');
	const index = join(built, 'index.html');
	if (!existsSync(index) || statSync(index).size === 0) {
		throw new Error(`${name}: the page build did not produce ${index}`);
	}
	cpSync(built, join(out, name), { recursive: true });
}

console.log(`Built the site into ${out}: launcher and ${selected.join(', ')}.`);
