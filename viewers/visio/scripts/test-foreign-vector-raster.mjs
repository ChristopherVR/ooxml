import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';
import { sanitizeVisioForeignVectorTree } from 'ooxml-core/visio';
import { renderForeignVector } from '../dist/render-foreign-vector.js';
const requireCore = createRequire(new URL('../../ooxml/package.json', import.meta.url));
const { createCanvas, loadImage } = requireCore('@napi-rs/canvas');
const dom = new JSDOM('<!doctype html>');
const shape = (d, extra = {}) => ({ tag: 'path', attrs: { d, ...extra } });
const source = (matrix) => ({
	tag: 'svg',
	attrs: { xmlns: 'http://www.w3.org/2000/svg', width: 30, height: 20, viewBox: '0 0 30 20' },
	children: [
		{
			tag: 'defs',
			attrs: {},
			children: [
				{ tag: 'clipPath', attrs: { id: 'left' }, children: [shape('M0 0H10V20H0Z')] },
				{
					tag: 'clipPath',
					attrs: { id: 'top', 'clip-path': 'url(#left)' },
					children: [shape('M0 0H30V10H0Z')],
				},
			],
		},
		{
			tag: 'g',
			attrs: { 'clip-path': 'url(#top)', ...(matrix ? { transform: matrix } : {}) },
			children: [shape('M-10 -10H40V30H-10Z', { fill: '#ff0000', stroke: 'none' })],
		},
	],
});
const comparisons = [];
try {
	for (const [matrix, inside, outside] of [
		[
			undefined,
			[5, 5],
			[
				[15, 5],
				[5, 15],
			],
		],
		[
			'matrix(-1 0 0 1 30 0)',
			[25, 5],
			[
				[15, 5],
				[25, 15],
			],
		],
	]) {
		const scene = sanitizeVisioForeignVectorTree(source(matrix));
		const svg = renderForeignVector(scene, dom.window.document);
		const serialized = new dom.window.XMLSerializer().serializeToString(svg);
		const bytes = execFileSync(
			'python3',
			[fileURLToPath(new URL('./rasterize-svg.py', import.meta.url))],
			{ input: serialized, maxBuffer: 1024 * 1024 },
		);
		const image = await loadImage(Buffer.from(serialized));
		assert.equal(image.width, 30);
		assert.equal(image.height, 20);
		const context = createCanvas(30, 20).getContext('2d');
		context.drawImage(image, 0, 0);
		const rgba = ([x, y]) => Array.from(context.getImageData(x, y, 1, 1).data);
		assert.deepEqual(rgba(inside), [255, 0, 0, 255]);
		for (const point of outside) assert.equal(rgba(point)[3], 0);
		const other = await loadImage(bytes),
			otherContext = createCanvas(30, 20).getContext('2d');
		otherContext.drawImage(other, 0, 0);
		const otherAlpha = outside.map(([x, y]) => otherContext.getImageData(x, y, 1, 1).data[3]);
		comparisons.push({
			matrix: matrix ?? 'identity',
			skiaExpectedPixels: true,
			librsvgOutsideAlpha: otherAlpha,
			librsvgNestedIntersection: otherAlpha.every((a) => a === 0),
		});
		if (process.argv.includes('--require-librsvg'))
			assert.ok(
				otherAlpha.every((a) => a === 0),
				'librsvg does not honor this nested clipPath intersection',
			);
	}
	console.log(
		'Secondary Skia SVG: generated path-only clip intersections, transforms and intrinsic viewport passed. Not browser or native Visio evidence.',
	);
	console.log(
		JSON.stringify({
			canvas: requireCore('@napi-rs/canvas/package.json').version,
			comparisons,
			note: 'The installed librsvg/Cairo backend ignores nested clipPath intersections; its divergent pixels are reported, not accepted as correct.',
		}),
	);
} finally {
	dom.window.close();
}
