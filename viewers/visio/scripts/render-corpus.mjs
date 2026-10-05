import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { resolve, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { JSDOM } from 'jsdom';
import { parseVsdx } from 'ooxml-core/visio';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const requireCore = createRequire(resolve(root, '..', 'ooxml', 'package.json'));
const native = requireCore('@napi-rs/canvas');
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const nativeEntry = requireCore.resolve('@napi-rs/canvas');
const nativePackage = JSON.parse(
	await readFile(resolve(dirname(nativeEntry), 'package.json'), 'utf8'),
);
const toolchain = {
	node: process.version,
	canvas: nativePackage.version,
	canvasEntrySha256: hash(await readFile(nativeEntry)),
	generator: 'scripts/render-corpus.mjs',
	generatorSha256: hash(await readFile(fileURLToPath(import.meta.url))),
	coreBundleSha256: hash(await readFile(resolve(root, '..', 'ooxml', 'dist/visio/index.mjs'))),
	viewerRendererSha256: hash(await readFile(resolve(root, 'dist/render-svg.js'))),
};
const input = resolve(process.argv[2] ?? ''),
	output = resolve(process.argv[3] ?? 'corpus-renders');
if (!process.argv[2])
	throw new Error('Usage: node scripts/render-corpus.mjs INPUT_DIRECTORY OUTPUT_DIRECTORY');
await mkdir(output, { recursive: true });
const dom = new JSDOM('<!doctype html>');
for (const key of ['document', 'HTMLElement', 'customElements', 'SVGElement', 'XMLSerializer'])
	globalThis[key] = dom.window[key];
const measurement = native.createCanvas(1, 1).getContext('2d');
globalThis.CanvasRenderingContext2D = measurement.constructor;
dom.window.HTMLCanvasElement.prototype.getContext = () => measurement;
const OriginalBlob = globalThis.Blob;
globalThis.Blob = class extends OriginalBlob {
	constructor(parts, options) {
		super(parts, options);
		this.renderBytes = Buffer.concat(parts.map((part) => Buffer.from(part)));
	}
};
URL.createObjectURL = (blob) => `data:${blob.type};base64,${blob.renderBytes.toString('base64')}`;
URL.revokeObjectURL = () => {};
const { renderPage } = await import('../dist/render-svg.js');
const results = [];
for (const filename of (await readdir(input)).filter((name) => name.endsWith('.vsdx')).sort()) {
	try {
		const fixture = await readFile(resolve(input, filename));
		const fixtureSha256 = hash(fixture);
		const document = await parseVsdx(fixture);
		for (const [index, page] of document.pages.entries()) {
			const result = renderPage(document, page),
				svg = result.svg;
			const scale = Math.min(96, 2048 / Math.max(page.width, page.height));
			const width = Math.max(1, Math.round(page.width * scale)),
				height = Math.max(1, Math.round(page.height * scale));
			svg.setAttribute('width', String(width));
			svg.setAttribute('height', String(height));
			const xml = new XMLSerializer().serializeToString(svg);
			const stem = `${basename(filename, '.vsdx')}-page${index + 1}`;
			await writeFile(resolve(output, `${stem}.svg`), xml);
			const bitmap = await native.loadImage(Buffer.from(xml));
			const canvas = native.createCanvas(width, height),
				context = canvas.getContext('2d');
			context.drawImage(bitmap, 0, 0, width, height);
			const png = canvas.toBuffer('image/png');
			await writeFile(resolve(output, `${stem}.png`), png);
			const pixels = context.getImageData(0, 0, width, height).data;
			let minX = width,
				minY = height,
				maxX = -1,
				maxY = -1;
			for (let y = 0; y < height; y++)
				for (let x = 0; x < width; x++)
					if (pixels[(y * width + x) * 4 + 3] > 16) {
						minX = Math.min(minX, x);
						minY = Math.min(minY, y);
						maxX = Math.max(maxX, x);
						maxY = Math.max(maxY, y);
					}
			if (maxX >= minX) {
				const crop = native.createCanvas(maxX - minX + 1, maxY - minY + 1);
				crop
					.getContext('2d')
					.drawImage(canvas, minX, minY, crop.width, crop.height, 0, 0, crop.width, crop.height);
				await writeFile(resolve(output, `${stem}-content.png`), crop.toBuffer('image/png'));
			}
			results.push({
				file: filename,
				fixtureSha256,
				svgSha256: hash(xml),
				pngSha256: hash(png),
				pixelsPerInch: scale,
				pageWidthInches: page.width,
				pageHeightInches: page.height,
				pageId: page.id,
				pageName: page.name,
				width,
				height,
				contentBounds:
					maxX >= minX
						? { x: minX, y: minY, width: maxX - minX + 1, height: maxY - minY + 1 }
						: null,
				diagnostics: document.diagnostics,
				renderWarnings: result.warnings,
			});
			result.dispose();
		}
	} catch (error) {
		results.push({ file: filename, error: error.message });
	}
}
await writeFile(
	resolve(output, 'manifest.json'),
	JSON.stringify(
		{
			toolchain,
			options: {
				maximumPageDimensionPx: 2048,
				maximumPixelsPerInch: 96,
				contentCropAlphaThreshold: 16,
				fontsPinned: false,
			},
			method:
				'Current SVG DOM renderer + @napi-rs/canvas rasterization and font measurement; not browser or Visio certification',
			results,
		},
		null,
		2,
	),
);
console.log(
	JSON.stringify(
		results.map(({ file, pageId, error, contentBounds }) => ({
			file,
			pageId,
			error,
			contentBounds,
		})),
		null,
		2,
	),
);
