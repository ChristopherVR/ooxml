import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import { rasterizeSvgInNode } from './index';

const SVG = new TextEncoder().encode(
	'<svg xmlns="http://www.w3.org/2000/svg" width="100" height="60"><rect width="100" height="60" fill="red"/></svg>',
);

/** Width and height from a PNG's IHDR chunk. */
function pngSize(png: Uint8Array): [number, number] {
	const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
	return [view.getUint32(16), view.getUint32(20)];
}

describe('rasterizeSvgInNode', () => {
	it('rasterises an SVG to a PNG of the size the writer asks for', async () => {
		const asked: number[][] = [];
		const png = await rasterizeSvgInNode(SVG, (w, h) => {
			asked.push([w, h]);
			return { width: w * 2, height: h * 2 };
		});
		expect(asked).toStrictEqual([[100, 60]]);
		expect(Array.from(png!.subarray(1, 4))).toStrictEqual([0x50, 0x4e, 0x47]); // "PNG"
		expect(pngSize(png!)).toStrictEqual([200, 120]);
	});

	it('returns nothing for bytes that are not an SVG', async () => {
		const png = await rasterizeSvgInNode(new Uint8Array([1, 2, 3]), () => ({
			width: 1,
			height: 1,
		}));
		expect(png).toBeUndefined();
	});
});

describe('the Node-only entry', () => {
	it('is the only part of the pptx area that names the Node canvas package', () => {
		const pptxRoot = path.resolve(__dirname, '..');
		const offenders = (readdirSync(pptxRoot, { recursive: true }) as string[])
			.map((file) => file.split(path.sep).join('/'))
			.filter((file) => /\.ts$/.test(file) && !/(?:^|\/)__tests__\/|\.test\.ts$/.test(file))
			.filter((file) => !file.startsWith('node/'))
			.filter((file) =>
				readFileSync(path.join(pptxRoot, file), 'utf8').includes('@napi-rs/canvas'),
			);
		expect(offenders).toStrictEqual([]);
	});
});
