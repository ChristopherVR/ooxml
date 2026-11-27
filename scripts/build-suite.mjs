import { cp, mkdir, rm } from 'node:fs/promises';
import { dirname, resolve, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { bundleSuite } from './bundle-suite.mjs';
import { buildPwa } from './build-pwa.mjs';
import { writeReleasePages } from './build-release-notes.mjs';
import { writeSamplesPage } from './build-samples-page.mjs';
import { applyContentSecurityPolicy, assertContentSecurityPolicy } from './csp.mjs';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/**
 * The suite site (shell, bundle and installable apps) in `out`, `suite-dist` by default.
 * `samples: false` (`--no-samples` on the command line) leaves the samples gallery and its files
 * out: the desktop app is built that way, so its installer does not carry them.
 */
export async function buildSuite(
	out = join(root, 'suite-dist'),
	{ logLevel, samples = true } = {},
) {
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
	// The release posts (docs/releases/*.md) at releases/<name>/, styled by site/releases/.
	writeReleasePages(out);
	// The samples gallery (samples/manifest.json) at samples/, with the files it offers. An
	// earlier build's copy must not survive a build without it.
	if (samples) writeSamplesPage(out);
	else await rm(join(out, 'samples'), { recursive: true, force: true });
	await bundleSuite(join(out, 'suite.js'), { logLevel });
	// The policy goes in before the PWA build, which copies the shell and hashes every file.
	applyContentSecurityPolicy(out);
	await buildPwa(out);
	assertContentSecurityPolicy(out);
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href)
	await buildSuite(undefined, { samples: !process.argv.includes('--no-samples') });
