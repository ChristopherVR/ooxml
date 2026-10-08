import { cp, mkdir } from 'node:fs/promises';
import { dirname, resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { bundleSuite } from './bundle-suite.mjs';
import { buildPwa } from './build-pwa.mjs';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const out = join(root, 'suite-dist');
await mkdir(out, { recursive: true });
for (const name of [
	'index.html',
	'styles.css',
	'themes.css',
	'workspace.css',
	'workspace-chrome.css',
	'workspace-library.css',
	'workspace-account.css',
	'workspace-files.css',
	'profile-form.css',
	'workspace-responsive.css',
	'favicon.svg',
	'appearance-init.js',
	'start-init.js',
])
	await cp(join(root, 'site', name), join(out, name));
await bundleSuite(join(out, 'suite.js'));
await buildPwa(out);
