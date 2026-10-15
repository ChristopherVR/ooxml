/** Serve the built Office application, including standalone PWA routes. */
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { dirname, extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SITE = resolve(ROOT, process.argv[3] ?? 'suite-dist');
const port = Number(process.argv[2] ?? 8123);
const TYPES = {
	'.html': 'text/html; charset=utf-8',
	'.js': 'text/javascript; charset=utf-8',
	'.css': 'text/css; charset=utf-8',
	'.svg': 'image/svg+xml',
	'.png': 'image/png',
	'.webmanifest': 'application/manifest+json',
};

const server = createServer(async (req, res) => {
	const path = decodeURIComponent(new URL(req.url ?? '/', 'http://x').pathname);
	const file = resolve(SITE, `.${path.endsWith('/') ? path + 'index.html' : path}`);
	if (!file.startsWith(SITE + '/') && !file.startsWith(SITE + '\\'))
		return res.writeHead(403).end();
	try {
		// Read before writing the head, so a missing file is a clean 404 and not a crash. The docs
		// sites use clean URLs (/getting-started for getting-started.html), so try that second.
		const body = await readFile(file).catch((error) =>
			!extname(file) ? readFile(`${file}.html`) : Promise.reject(error),
		);
		const type =
			TYPES[extname(file)] ?? (extname(file) ? 'application/octet-stream' : TYPES['.html']);
		res.writeHead(200, { 'content-type': type });
		res.end(body);
	} catch {
		res.writeHead(404).end('Not found');
	}
});
server.listen(port, '127.0.0.1', () =>
	console.log(`Suite on http://localhost:${server.address().port}/`),
);
