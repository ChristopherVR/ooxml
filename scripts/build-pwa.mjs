import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { join, posix } from 'node:path';
import { createHash } from 'node:crypto';
import { deflateSync } from 'node:zlib';
import { bundleShellCss } from './shell-css.mjs';

const products = [
	['office', 'OOXML Office', '', '#0f6cbd'],
	['word', 'OOXML Word', 'docx', '#185abd'],
	['excel', 'OOXML Excel', 'xlsx', '#107c41'],
	['powerpoint', 'OOXML PowerPoint', 'pptx', '#c43e1c'],
	['visio', 'OOXML Visio', 'vsdx', '#3955a3'],
	['teams', 'OOXML Teams', 'teams', '#6264a7'],
];
function icon(size, color) {
	const rgb = color.match(/\w\w/g).map((v) => parseInt(v, 16));
	const pixels = Buffer.alloc((size * 4 + 1) * size);
	for (let y = 0; y < size; y++)
		for (let x = 0; x < size; x++) {
			const i = y * (size * 4 + 1) + 1 + x * 4;
			const tile =
				x > size * 0.23 &&
				x < size * 0.77 &&
				y > size * 0.23 &&
				y < size * 0.77 &&
				Math.abs(x - size / 2) > size * 0.035 &&
				Math.abs(y - size / 2) > size * 0.035;
			pixels.set(tile ? [255, 255, 255, 255] : [...rgb, 255], i);
		}
	function chunk(type, data) {
		const body = Buffer.concat([Buffer.from(type), data]);
		let crc = 0xffffffff;
		for (const byte of body) {
			crc ^= byte;
			for (let b = 0; b < 8; b++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
		}
		const len = Buffer.alloc(4),
			end = Buffer.alloc(4);
		len.writeUInt32BE(data.length);
		end.writeUInt32BE((crc ^ 0xffffffff) >>> 0);
		return Buffer.concat([len, body, end]);
	}
	const header = Buffer.alloc(13);
	header.writeUInt32BE(size, 0);
	header.writeUInt32BE(size, 4);
	header[8] = 8;
	header[9] = 6;
	return Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', deflateSync(pixels)),
		chunk('IEND', Buffer.alloc(0)),
	]);
}
export async function buildPwa(out) {
	const shellCss = await bundleShellCss(out);
	const template = await readFile(join(out, 'index.html'), 'utf8');
	await mkdir(join(out, 'icons'), { recursive: true });
	for (const [slug, name, product, color] of products) {
		const prefix = product ? '../../' : './';
		const dir = product ? join(out, 'apps', slug) : out;
		await mkdir(dir, { recursive: true });
		for (const size of [192, 512])
			await writeFile(join(out, 'icons', `${slug}-${size}.png`), icon(size, color));
		const manifest = {
			id: './',
			name,
			short_name: product ? name.replace('OOXML ', '') : 'Office',
			start_url: './',
			scope: './',
			display: 'standalone',
			background_color: '#ffffff',
			theme_color: color,
			// Opening a OneDrive link in the installed app (apps/office-suite/src/launch-link.js): a
			// link shared to the app, or a web+ooxml:<link> address (the suite only).
			share_target: {
				action: './',
				method: 'GET',
				params: { title: 'title', text: 'text', url: 'open' },
			},
			...(product ? {} : { protocol_handlers: [{ protocol: 'web+ooxml', url: './?open=%s' }] }),
			icons: [192, 512].map((size) => ({
				src: `${prefix}icons/${slug}-${size}.png`,
				sizes: `${size}x${size}`,
				type: 'image/png',
				purpose: 'any maskable',
			})),
		};
		await writeFile(join(dir, 'manifest.webmanifest'), JSON.stringify(manifest, null, 2));
		if (product) {
			const html = template
				.replace('<body>', `<body data-product="${product}">`)
				.replace('<title>OOXML Office</title>', `<title>${name}</title>`)
				.replace('name="ooxml-root" content="./"', 'name="ooxml-root" content="../../"')
				.replace(
					/(src|href)="(appearance-init\.js|favicon\.svg|shell-[0-9a-f]+\.css|suite\.js)"/g,
					`$1="../../$2"`,
				)
				.replace('href="icons/office-192.png"', `href="../../icons/${slug}-192.png"`);
			await writeFile(join(dir, 'index.html'), html);
		}
	}
	const files = [];
	async function walk(dir, prefix = '') {
		for (const entry of await readdir(dir, { withFileTypes: true })) {
			const rel = prefix + entry.name;
			if (
				(entry.isDirectory() && ['suite-assets', 'apps', 'icons'].includes(entry.name)) ||
				(entry.isDirectory() && prefix)
			)
				await walk(join(dir, entry.name), rel + '/');
			else if (
				entry.isFile() &&
				/\.(js|css|html|png|svg|webmanifest)$/.test(entry.name) &&
				entry.name !== 'sw.js' &&
				// The not-found page is not the app shell (the worker refuses to cache one that is
				// not), and it is only ever served by the host for a missing address.
				!/^404\./.test(rel)
			)
				files.push(rel);
		}
	}
	await walk(out);
	files.sort();
	// The shell's scripts keep fixed names, so every reference to one from a page gets
	// ?v=<hash of every asset>: a deploy changes the URL, which bypasses the HTTP cache and any
	// older service worker. The shell stylesheet and the chunks in suite-assets are content-hashed.
	const assets = createHash('sha256');
	for (const file of files.filter((f) => !f.endsWith('.html')))
		assets.update(file).update(await readFile(join(out, file)));
	const stamp = assets.digest('hex').slice(0, 12);
	const stamped = new Set();
	const versioned = (text, pattern) =>
		text.replace(pattern, (match, before, path, name, after) => {
			if (!files.includes(name) || name === shellCss) return match;
			stamped.add(name);
			return `${before}${path}${name}?v=${stamp}${after}`;
		});
	for (const file of files.filter((f) => f.endsWith('.html'))) {
		const text = await readFile(join(out, file), 'utf8');
		await writeFile(
			join(out, file),
			versioned(text, /((?:src|href)=")((?:\.\.\/\.\.\/)?)([\w-]+\.(?:js|css))(")/g)
				// The app shows this as its version (apps/office-suite/src/updates.js).
				.replace(
					'<meta name="ooxml-root"',
					`<meta name="ooxml-build" content="${stamp}" />
		<meta name="ooxml-root"`,
				),
		);
	}
	// Subresource integrity on the shell's own tags, so the browser refuses a script or style that
	// is not this build's, whoever serves it: the stamped scripts and the content-hashed shell sheet.
	for (const file of files.filter((f) => f.endsWith('.html'))) {
		const text = await readFile(join(out, file), 'utf8');
		let next = text;
		for (const [tag, src] of text.matchAll(
			/<(?:script\b|link\b(?=[^>]*\brel="stylesheet"))[^>]*?(?:src|href)="([^"?]+)(?:\?v=\w+)?"[^>]*>/g,
		)) {
			const path = posix.join(posix.dirname(file), src);
			if (!files.includes(path) || /\sintegrity=/.test(tag)) continue;
			const digest = createHash('sha384')
				.update(await readFile(join(out, path)))
				.digest('base64');
			next = next.replace(
				tag,
				tag.replace(/\s*\/?>$/, (end) => ` integrity="sha384-${digest}"${end}`),
			);
		}
		if (next !== text) await writeFile(join(out, file), next);
	}
	// The worker serves a saved copy only when it still has the SHA-256 recorded here, because any
	// page of the origin can rewrite Cache Storage (scripts/pwa/service-worker.js). The version is
	// the hash of the table itself, so a later worker can check an older build's saved table.
	const table = {};
	for (const file of files)
		table[stamped.has(file) ? `${file}?v=${stamp}` : file] = createHash('sha256')
			.update(await readFile(join(out, file)))
			.digest('base64');
	const json = JSON.stringify(table);
	const version = createHash('sha256').update(json).digest('hex');
	const worker = await readFile(new URL('./pwa/service-worker.js', import.meta.url), 'utf8');
	await writeFile(
		join(out, 'sw.js'),
		`const VERSION=${JSON.stringify(version)};\nconst FILES=${json};\n${worker}`,
	);
}
