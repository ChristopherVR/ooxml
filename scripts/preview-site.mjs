/**
 * Serve the suite, or the full Pages build with its published URL prefix:
 * node scripts/preview-site.mjs 8123 pages-dist /ooxml/
 */
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { dirname, extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SITE = resolve(ROOT, process.argv[3] ?? 'suite-dist');
const port = Number(process.argv[2] ?? 8123);
const base = '/' + (process.argv[4] ?? '').split('/').filter(Boolean).join('/');
const prefix = base === '/' ? '' : base;
// Full Pages previews include the docs; suite-only package previews still need the published guides.
// Rewrite compiled links too: the docs generate demo iframe and full-screen URLs in JavaScript.
const localize = (body, type) =>
	prefix && /^(text\/|application\/(json|manifest\+json))/.test(type)
		? body.toString().replaceAll('https://christophervr.github.io/ooxml/', `${prefix}/`)
		: body;
const TYPES = {
	'.html': 'text/html; charset=utf-8',
	'.js': 'text/javascript; charset=utf-8',
	'.css': 'text/css; charset=utf-8',
	'.svg': 'image/svg+xml',
	'.png': 'image/png',
	'.webmanifest': 'application/manifest+json',
	'.json': 'application/json',
};

const server = createServer(async (req, res) => {
	const url = new URL(req.url ?? '/', 'http://x');
	if (prefix && !url.pathname.startsWith(prefix + '/')) {
		const path = url.pathname === prefix ? '/' : url.pathname;
		res.writeHead(302, { location: prefix + path + url.search }).end();
		return;
	}
	const path = decodeURIComponent(url.pathname.slice(prefix.length));
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
		res.end(localize(body, type));
	} catch {
		// Like GitHub Pages, answer a missing path with the site's own 404.html: the one at the
		// served root, or at the first path segment when the root mirrors a project site (/ooxml/).
		const [first] = path.split('/').filter(Boolean);
		for (const dir of [first ? resolve(SITE, first) : SITE, SITE]) {
			try {
				const body = await readFile(resolve(dir, '404.html'));
				res.writeHead(404, { 'content-type': TYPES['.html'] });
				return res.end(localize(body, TYPES['.html']));
			} catch {}
		}
		res.writeHead(404).end('Not found');
	}
});
server.listen(port, '127.0.0.1', () =>
	console.log(`Site on http://localhost:${server.address().port}${prefix}/`),
);
