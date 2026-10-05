// Builds the GitHub Pages site: the VitePress documentation in docs/, then each demo into its own
// route under docs/.vitepress/dist. Adapted from ChristopherVR/docx-viewer scripts/build-pages.mjs.
//
// Pages is static, so there is no server behind the demos. They are built with VITE_TEAMS_STATIC=1,
// which starts them in the core's local mode (tabs of one browser share state over
// BroadcastChannel) and shows a notice saying so. Nothing pretends to be a hosted backend.
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, rmSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const docs = resolve(root, 'docs');
const dist = resolve(docs, '.vitepress', 'dist');
const node = process.execPath;
const vitepressCli = resolve(docs, 'node_modules', 'vitepress', 'bin', 'vitepress.js');
const viteCli = resolve(root, 'node_modules', 'vite', 'bin', 'vite.js');
const base = '/teams-viewer/';
// `/demo/` is the long-standing route for the primary demo in every viewer's Pages site.
const demos = [
	{ dir: 'vanilla', route: 'demo' },
	{ dir: 'react', route: 'demo-react' },
];

function run(script, args, cwd, env = process.env) {
	const result = spawnSync(node, [script, ...args], { cwd, env, stdio: 'inherit' });
	if (result.error) throw result.error;
	if (result.status !== 0)
		throw new Error(`Command failed (${result.status}): ${script} ${args.join(' ')}`);
}

if (!existsSync(vitepressCli)) {
	throw new Error(
		'VitePress is missing. Install the docs dependencies with `bun install --cwd docs`.',
	);
}
if (!existsSync(viteCli)) {
	throw new Error('Vite is missing. Install the workspace dependencies with `bun install`.');
}

run(vitepressCli, ['build', '.'], docs);

for (const { dir, route } of demos) {
	const outDir = resolve(dist, route);
	if (existsSync(outDir)) rmSync(outDir, { recursive: true, force: true });
	mkdirSync(outDir, { recursive: true });
	run(
		viteCli,
		['build', '--base', `${base}${route}/`, '--outDir', outDir, '--emptyOutDir'],
		resolve(root, 'demos', dir),
		{ ...process.env, VITE_TEAMS_STATIC: '1' },
	);
	const index = resolve(outDir, 'index.html');
	if (!existsSync(index) || statSync(index).size === 0) {
		throw new Error(`The ${dir} demo build did not produce ${index}`);
	}
}

console.log(
	`Built the documentation and ${demos.length} demos: ${demos.map((d) => `${base}${d.route}/`).join(', ')}`,
);
