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
import { buildPwa } from './build-pwa.mjs';
import { writeReleasePages } from './build-release-notes.mjs';
import { writeSamplesPage } from './build-samples-page.mjs';
import { applyContentSecurityPolicy, assertContentSecurityPolicy } from './csp.mjs';

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
/** The URL path the site is served from. */
const SITE_BASE = '/ooxml/';

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
// The release posts (docs/releases/*.md) at releases/<name>/; this also drops the page template.
writeReleasePages(out);
// The samples gallery at samples/, with the sample files and thumbnails; this drops its template too.
writeSamplesPage(out);
await bundleSuite(join(out, 'suite.js'));
// The policy goes in before the PWA build, which copies the shell and hashes every file.
applyContentSecurityPolicy(out, { base: SITE_BASE });
await buildPwa(out);

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

// The origin is shared with every other project site on christophervr.github.io, so every page,
// the viewers' docs and demos included, carries the suite's script policy (scripts/csp.mjs). Pages
// already carrying it, the suite's among them, are left as they are.
applyContentSecurityPolicy(out, { base: SITE_BASE });
assertContentSecurityPolicy(out);

console.log(`Built the site into ${out}: launcher and ${selected.join(', ')}.`);
