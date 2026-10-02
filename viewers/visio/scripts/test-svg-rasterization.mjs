import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

// Secondary SVG renderer evidence only. This is not browser or Microsoft Visio visual parity.
// Uses the native canvas already installed by the private sibling core checkout.
const requireCore = createRequire(new URL('../../ooxml/package.json', import.meta.url));
const { createCanvas, loadImage } = requireCore('@napi-rs/canvas');
const dom = new JSDOM('<!doctype html>');
globalThis.document = dom.window.document;
globalThis.XMLSerializer = dom.window.XMLSerializer;
const { exportPageSvg } = await import('../dist/export-svg.js');
const { demoDocument } = await import('../dist/demo-document.js');

try {
	const source = createCanvas(32, 32),
		context = source.getContext('2d');
	for (const [x, y, color] of [
		[0, 0, '#ff0000'],
		[16, 0, '#00ff00'],
		[0, 16, '#0000ff'],
		[16, 16, '#ffff00'],
	]) {
		context.fillStyle = color;
		context.fillRect(x, y, 16, 16);
	}
	for (const format of ['png', 'jpeg', 'gif']) {
		const bytes = Uint8Array.from(await source.encode(format, 100));
		const model = structuredClone(demoDocument),
			page = model.pages[0],
			shape = page.shapes[0];
		model.pages = [page];
		page.width = 4;
		page.height = 3;
		shape.width = 2;
		shape.height = 2;
		shape.transform = [1, 0, 0, 1, 1, 0.5];
		shape.geometry = [];
		shape.text.plainText = '';
		shape.text.runs = [];
		shape.image = {
			mimeType: `image/${format}`,
			bytes,
			pixelWidth: 32,
			pixelHeight: 32,
			x: -0.5,
			y: 0.5,
			width: 2,
			height: 2,
		};
		// A second transformed instance shares the same encoded symbol with independent opacity.
		const copy = {
			...shape,
			id: 'copy',
			transform: [1, 0, 0, 1, 3, 0.5],
			image: { ...shape.image, opacity: 0.5 },
		};
		page.shapes = [shape, copy];
		const output = exportPageSvg(model);
		assert.equal(
			(output.svg.match(/data:image\//g) ?? []).length,
			1,
			`${format}: encoded asset deduplication`,
		);
		assert.equal(
			(output.svg.match(/<use /g) ?? []).length,
			2,
			`${format}: two independently placed instances`,
		);
		// Native canvas' Skia SVG subset omits symbols and embedded raster images. Use installed
		// librsvg/Cairo for actual SVG rendering, then native canvas only for PNG pixel inspection.
		const raster = execFileSync(
			'python3',
			[fileURLToPath(new URL('./rasterize-svg.py', import.meta.url))],
			{ input: output.svg, maxBuffer: 8 * 1024 * 1024 },
		);
		const image = await loadImage(raster);
		const dpi = 96;
		assert.equal(image.width, page.width * dpi, `${format}: physical width`);
		assert.equal(image.height, page.height * dpi, `${format}: physical height`);
		const canvas = createCanvas(image.width, image.height),
			drawing = canvas.getContext('2d');
		drawing.drawImage(image, 0, 0);
		const rgba = (x, y) =>
			Array.from(drawing.getImageData(Math.round(x * dpi), Math.round(y * dpi), 1, 1).data);
		const expected = [
			[1.25, 0.75, [255, 0, 0, 255]],
			[2, 0.75, [0, 255, 0, 255]],
			[1.25, 1.5, [0, 0, 255, 255]],
			[2, 1.5, [255, 255, 0, 255]],
		];
		for (const [x, y, channels] of expected) {
			const actual = rgba(x, y);
			assert.ok(
				actual.every((value, index) => Math.abs(value - channels[index]) <= 8),
				`${format}: correct crop/flip color at ${x},${y}: ${actual}`,
			);
		}
		for (const [x, y] of [
			[0.75, 0.75],
			[1.25, 0.25],
			[1.25, 2.25],
			[2.75, 1.5],
		])
			assert.equal(
				rgba(x, y)[3],
				0,
				`${format}: outside clip or image is transparent at ${x},${y}`,
			);
		assert.ok(
			Math.abs(rgba(3.25, 0.75)[3] - 128) <= 1,
			`${format}: opacity belongs to each use instance`,
		);
		assert.ok(rgba(3.25, 0.75)[0] > 245, `${format}: repeated symbol has the same orientation`);
		console.log(
			`Secondary librsvg/Cairo rasterizer: ${format.toUpperCase()} dimensions, crop, flip, opacity and deduplicated reuse passed.`,
		);
	}
} finally {
	dom.window.close();
	delete globalThis.document;
	delete globalThis.XMLSerializer;
}
