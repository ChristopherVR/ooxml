/** Serve the built Office application, including standalone PWA routes. */
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SITE = join(ROOT, 'suite-dist');
const port = Number(process.argv[2] ?? 8123);
const TYPES = {
	'.html': 'text/html; charset=utf-8',
	'.js': 'text/javascript; charset=utf-8',
	'.css': 'text/css; charset=utf-8',
	'.svg': 'image/svg+xml',
	'.png': 'image/png',
	'.webmanifest': 'application/manifest+json',
};

createServer(async (req, res) => {
	const path = decodeURIComponent(new URL(req.url ?? '/', 'http://x').pathname);
	const file = resolve(SITE, `.${path.endsWith('/') ? path + 'index.html' : path}`);
	if (!file.startsWith(SITE + '/') && !file.startsWith(SITE + '\\'))
		return res.writeHead(403).end();
	try {
		// Read before writing the head, so a missing file is a clean 404 and not a crash.
		const body = await readFile(file);
		res.writeHead(200, { 'content-type': TYPES[extname(file)] ?? 'application/octet-stream' });
		res.end(body);
	} catch {
		res.writeHead(404).end('Not found');
	}
}).listen(port, '127.0.0.1', () => console.log(`Suite on http://localhost:${port}/`));
