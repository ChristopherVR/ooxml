import { cp, mkdir } from 'node:fs/promises';
import { dirname, resolve, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { bundleSuite } from './bundle-suite.mjs';
import { buildPwa } from './build-pwa.mjs';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** The suite site (shell, bundle and installable apps) in `out`, `suite-dist` by default. */
export async function buildSuite(out = join(root, 'suite-dist'), { logLevel } = {}) {
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
		'workspace-chat.css',
		'workspace-settings.css',
		'workspace-simple.css',
		'ai-config.json',
		'profile-form.css',
		'workspace-responsive.css',
		'workspace-mobile.css',
		'favicon.svg',
		'appearance-init.js',
		'404.html',
		'404.js',
	])
		await cp(join(root, 'site', name), join(out, name));
	await cp(join(root, 'site', 'docs'), join(out, 'docs'), { recursive: true });
	await bundleSuite(join(out, 'suite.js'), { logLevel });
	await buildPwa(out);
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) await buildSuite();
