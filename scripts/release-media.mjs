/**
 * release-media.mjs: where a release post's pictures and recordings come from, and how big they
 * are, so the page can copy them next to itself and reserve their space before they load.
 */
import { existsSync, readFileSync, statSync } from 'node:fs';
import { basename, dirname, extname, join, posix, relative, resolve } from 'node:path';

const REPO_BLOB = 'https://github.com/ChristopherVR/ooxml/blob/main/';
export const IMAGE = /\.(?:png|jpe?g|gif|webp|avif|svg)$/i;

/** Width and height of a PNG, GIF, JPEG, WebP or SVG file, or null when it cannot be read. */
export function imageSize(bytes) {
	const b = Buffer.from(bytes);
	if (b.length >= 24 && b.readUInt32BE(0) === 0x89504e47)
		return { width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
	if (b.length >= 10 && b.toString('ascii', 0, 3) === 'GIF')
		return { width: b.readUInt16LE(6), height: b.readUInt16LE(8) };
	if (
		b.length >= 30 &&
		b.toString('ascii', 0, 4) === 'RIFF' &&
		b.toString('ascii', 8, 12) === 'WEBP'
	) {
		const chunk = b.toString('ascii', 12, 16);
		if (chunk === 'VP8X')
			return { width: 1 + b.readUIntLE(24, 3), height: 1 + b.readUIntLE(27, 3) };
		if (chunk === 'VP8L') {
			const bits = b.readUInt32LE(21);
			return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 };
		}
		if (chunk === 'VP8 ')
			return { width: b.readUInt16LE(26) & 0x3fff, height: b.readUInt16LE(28) & 0x3fff };
	}
	if (b.length >= 4 && b[0] === 0xff && b[1] === 0xd8) {
		let i = 2;
		while (i + 9 < b.length) {
			if (b[i] !== 0xff) return null;
			const marker = b[i + 1];
			if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker))
				return { width: b.readUInt16BE(i + 7), height: b.readUInt16BE(i + 5) };
			i += 2 + b.readUInt16BE(i + 2);
		}
		return null;
	}
	const text = b.toString('utf8', 0, Math.min(b.length, 2048));
	const svg = /<svg\b[^>]*>/i.exec(text)?.[0];
	if (svg) {
		const attr = (name) =>
			Number.parseFloat(new RegExp(`\\s${name}="([\\d.]+)(?:px)?"`).exec(svg)?.[1]);
		const box = /\sviewBox="[\d.-]+[\s,]+[\d.-]+[\s,]+([\d.]+)[\s,]+([\d.]+)"/.exec(svg);
		const width = attr('width') || Number(box?.[1]);
		const height = attr('height') || Number(box?.[2]);
		if (width && height) return { width: Math.round(width), height: Math.round(height) };
	}
	return null;
}

/**
 * URL resolution for the post `<postsDir>/<name>.md`, served from `releases/<name>/`. A picture
 * in the repository is copied to `assets/` beside the page (`assets` maps the page path to the
 * source file); `assets/x.gif` is looked for in `<postsDir>/<name>/assets/` when it is not beside
 * the post. Any other repository path links to GitHub; absolute URLs and anchors stay.
 */
export function mediaResolver({ root, postsDir, name, assets }) {
	const find = (path) => {
		const direct = resolve(postsDir, path);
		if (existsSync(direct)) return direct;
		const own = resolve(postsDir, name, path);
		return existsSync(own) ? own : direct;
	};
	const copy = (file) => {
		const page = `assets/${basename(file)}`;
		const known = assets.get(page);
		if (known && known !== file) throw new Error(`Two pictures are named ${basename(file)}`);
		assets.set(page, file);
		return page;
	};
	const link = (url, kind) => {
		if (/^(?:[a-z][\w+.-]*:|#|\/\/)/i.test(url)) return url;
		const [path, hash = ''] = url.split('#');
		const file = find(decodeURIComponent(path));
		const repoPath = relative(root, file).split('\\').join('/');
		if (repoPath.startsWith('..')) return url;
		if (kind === 'img' || IMAGE.test(path)) return copy(file);
		const isDir = existsSync(file) && statSync(file).isDirectory();
		const base = isDir ? REPO_BLOB.replace('/blob/', '/tree/') : REPO_BLOB;
		return `${base}${posix.normalize(repoPath)}${hash ? `#${hash}` : ''}`;
	};
	/** A figure's picture: its page URL, size, and the recording and poster beside a GIF. */
	const media = (url) => {
		if (/^(?:[a-z][\w+.-]*:|\/\/)/i.test(url)) return { src: url };
		const file = find(decodeURIComponent(url));
		if (!existsSync(file)) throw new Error(`The picture ${url} does not exist`);
		const info = { src: copy(file), ...imageSize(readFileSync(file)) };
		const stem = join(dirname(file), basename(file, extname(file)));
		if (extname(file).toLowerCase() === '.gif' && existsSync(`${stem}.mp4`)) {
			info.video = copy(`${stem}.mp4`);
			const poster = ['.poster.webp', '.poster.png', '.poster.jpg']
				.map((ext) => stem + ext)
				.find((path) => existsSync(path));
			if (poster) info.poster = copy(poster);
		}
		return info;
	};
	return { link, media };
}
