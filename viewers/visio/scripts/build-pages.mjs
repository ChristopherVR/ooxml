/**
 * Builds the GitHub Pages site: the VitePress documentation in docs/, then the playground into
 * /demo/ (vanilla; the route the Office launcher has always embedded), a copy of it at
 * /demo-vanilla/ (the name the other viewers use) and one demo per framework into /demo-<id>/.
 *
 * `node scripts/build-pages.mjs [outDir] [--browser-tests]`: docs/.vitepress/dist by default. The
 * browser-test build serves the site from "/" (DOCS_BASE) and adds the test API entry.
 */
import { spawnSync } from 'node:child_process';
import { cpSync, existsSync, rmSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { FRAMEWORK_DEMOS } from './framework-demos.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const browserTests = args.includes('--browser-tests');
const outArg = args.find((arg) => !arg.startsWith('--'));
const dist = resolve(root, outArg ?? 'docs/.vitepress/dist');
const node = process.execPath;
const vitepressCli = resolve(root, 'docs', 'node_modules', 'vitepress', 'bin', 'vitepress.js');
const viteCli = resolve(root, 'node_modules', 'vite', 'bin', 'vite.js');

function run(script, scriptArgs, env = {}) {
	const result = spawnSync(node, [script, ...scriptArgs], {
		cwd: root,
		env: { ...process.env, ...env },
		stdio: 'inherit',
	});
	if (result.error) throw result.error;
	if (result.status !== 0)
		throw new Error(`Command failed (${result.status}): ${script} ${scriptArgs.join(' ')}`);
}

if (!existsSync(vitepressCli)) {
	throw new Error(
		'VitePress is missing. Install the docs dependencies with `npm ci --prefix docs`.',
	);
}

// 1. The documentation. VitePress empties its output folder, so it goes first.
run(vitepressCli, ['build', 'docs'], {
	DOCS_OUT: dist,
	...(browserTests ? { DOCS_BASE: '/' } : {}),
});

// 2. The playground at /demo/ (Vite writes demo/index.html and assets/ beside the docs).
run(viteCli, ['build', ...(browserTests ? ['--mode', 'browser-tests'] : [])], { PAGES_OUT: dist });
if (!existsSync(resolve(dist, 'demo/index.html')))
	throw new Error('The playground build did not produce demo/index.html.');

// 3. The same playground under the standard name. Its assets are relative (../assets/), so a copy
//    beside it keeps working.
const vanilla = resolve(dist, 'demo-vanilla');
rmSync(vanilla, { recursive: true, force: true });
cpSync(resolve(dist, 'demo'), vanilla, { recursive: true });

// 4. One demo per framework.
run(resolve(root, 'scripts/build-demos.mjs'), [dist]);
for (const { id } of FRAMEWORK_DEMOS) {
	const index = resolve(dist, `demo-${id}`, 'index.html');
	if (!existsSync(index) || statSync(index).size === 0)
		throw new Error(`The ${id} demo build did not produce ${index}`);
}

console.log(
	`Built the documentation and every demo into ${dist}: /demo/, /demo-vanilla/, ${FRAMEWORK_DEMOS.map((d) => `/demo-${d.id}/`).join(', ')}`,
);
