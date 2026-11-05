/**
 * Serve the Office suite with the product demos' OneDrive sign-in, for local review:
 *
 *   bun run build:suite && bun run preview:suite:onedrive [port]
 *
 * It bundles demos/onedrive/suite-preview.ts (an app registration through MSAL, or a pasted Graph
 * Explorer token, in the same dialog as the product demos) into a copy of suite-dist, loads it
 * before the suite entry, and serves that copy. suite-dist itself is never changed, and nothing
 * here is published: a real host supplies its own sign-in (docs/onedrive-integration.md).
 * Register http://localhost:<port>/ as a single-page application redirect URI to use your own app
 * registration; the default port is 8123.
 */
import { cp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SOURCE = resolve(ROOT, 'suite-dist');
const OUT = resolve(ROOT, '.suite-preview-onedrive');
const port = process.argv[2] ?? '8123';

try {
	await stat(resolve(SOURCE, 'index.html'));
} catch {
	console.error('suite-dist is missing. Run bun run build:suite first.');
	process.exit(1);
}

await rm(OUT, { recursive: true, force: true });
await cp(SOURCE, OUT, { recursive: true });
// An ES module: the sign-in code reads import.meta, and modules run in document order, so this one
// (first in <head>) has set the options before the suite's own module entry reads them.
const built = await Bun.build({
	entrypoints: [resolve(ROOT, 'demos/onedrive/suite-preview.ts')],
	target: 'browser',
	format: 'esm',
	splitting: true,
	minify: true,
	outdir: resolve(OUT, 'demo-onedrive'),
	naming: { entry: 'index.js', chunk: '[name]-[hash].js' },
});
if (!built.success) {
	for (const log of built.logs) console.error(log);
	process.exit(1);
}
const index = resolve(OUT, 'index.html');
const html = await readFile(index, 'utf8');
await writeFile(
	index,
	html.replace('<head>', '<head><script type="module" src="demo-onedrive/index.js"></script>'),
);

console.log(`Office suite with demo OneDrive sign-in: http://localhost:${port}/`);
const server = spawn(process.execPath, [resolve(ROOT, 'scripts/preview-site.mjs'), port, OUT], {
	stdio: 'inherit',
});
server.on('exit', (code) => process.exit(code ?? 0));
