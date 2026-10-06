/**
 * Preview the launcher locally: bundle the suite module into `site/` (git-ignored) and serve
 * `site/` on http://localhost:8123 (or the first argument). Locally the embedded demos are
 * cross-origin; see AGENTS.md.
 *
 *   node scripts/preview-site.mjs [port]
 */
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { bundleSuite } from './bundle-suite.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SITE = join(ROOT, 'site');
const port = Number(process.argv[2] ?? 8123);
const TYPES = {
	'.html': 'text/html; charset=utf-8',
	'.js': 'text/javascript; charset=utf-8',
	'.css': 'text/css; charset=utf-8',
	'.svg': 'image/svg+xml',
};

await bundleSuite(join(SITE, 'suite.js'));

createServer(async (req, res) => {
	const path = decodeURIComponent(new URL(req.url ?? '/', 'http://x').pathname);
	const file = resolve(SITE, `.${path === '/' ? '/index.html' : path}`);
	if (!file.startsWith(SITE)) return res.writeHead(403).end();
	try {
		res.writeHead(200, { 'content-type': TYPES[extname(file)] ?? 'application/octet-stream' });
		res.end(await readFile(file));
	} catch {
		res.writeHead(404).end('Not found');
	}
}).listen(port, () => console.log(`Launcher on http://localhost:${port}/`));
