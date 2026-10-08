import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { deflateSync } from 'node:zlib';

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
					/(src|href)="(appearance-init\.js|favicon\.svg|styles\.css|themes\.css|workspace\.css|suite\.js)"/g,
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
				entry.name !== 'sw.js'
			)
				files.push(rel);
		}
	}
	await walk(out);
	const hash = createHash('sha256');
	for (const file of files.sort()) hash.update(file).update(await readFile(join(out, file)));
	const worker = await readFile(new URL('./pwa/service-worker.js', import.meta.url), 'utf8');
	await writeFile(
		join(out, 'sw.js'),
		`const VERSION=${JSON.stringify(hash.digest('hex').slice(0, 16))};\nconst FILES=${JSON.stringify(files)};\n${worker}`,
	);
}
