import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { blurredSilhouette, drawBlurredSilhouette } from './pdf-blur-mask';
import { flattenSegments } from './pdf-raster';
import { parsePath } from './pdf-path';
import { buildVectorPdf, PdfVectorPage } from './pdf-vector';

const square = parsePath('M 0 0 L 100 0 L 100 100 L 0 100 Z')!;
const black = { r: 0, g: 0, b: 0 };

/** The alpha at a point of the silhouette's box (silhouette units). */
function alphaAt(mask: NonNullable<ReturnType<typeof blurredSilhouette>>, x: number, y: number) {
	const { resource } = mask;
	const column = Math.floor(((x - mask.x) / mask.width) * resource.width);
	const row = Math.floor(((y - mask.y) / mask.height) * resource.height);
	return resource.alpha!.data[row * resource.width + column]! / 255;
}

describe('blurredSilhouette', () => {
	it('blurs a filled square like a Gaussian: half at the edge, solid inside, clear outside', () => {
		const mask = blurredSilhouette([{ segments: square, fill: 'nonzero' }], 6, black)!;
		// Three deviations of room around the square.
		expect(mask.x).toBeCloseTo(-18);
		expect(mask.y).toBeCloseTo(-18);
		expect(alphaAt(mask, 50, 50)).toBeCloseTo(1, 1);
		expect(alphaAt(mask, 0, 50)).toBeGreaterThan(0.4);
		expect(alphaAt(mask, 0, 50)).toBeLessThan(0.6);
		// One deviation outside: about 16 % (the Gaussian's tail), then nearly nothing.
		expect(alphaAt(mask, -6, 50)).toBeGreaterThan(0.08);
		expect(alphaAt(mask, -6, 50)).toBeLessThan(0.25);
		expect(alphaAt(mask, -17, 50)).toBeLessThan(0.02);
		// The fall-off is monotone across the edge.
		const profile = [-15, -10, -5, 0, 5, 10, 15].map((x) => alphaAt(mask, x, 50));
		for (let i = 1; i < profile.length; i++) expect(profile[i]!).toBeGreaterThan(profile[i - 1]!);
	});

	it('keeps the raster small whatever the drawing size', () => {
		const huge = parsePath('M 0 0 L 10000 0 L 10000 5000 L 0 5000 Z')!;
		const mask = blurredSilhouette([{ segments: huge, fill: 'nonzero' }], 1, black)!;
		expect(Math.max(mask.resource.width, mask.resource.height)).toBeLessThanOrEqual(385);
	});

	it('honours the fill rule, strokes and the part matrix', () => {
		const ring = parsePath('M 0 0 L 100 0 L 100 100 L 0 100 Z M 25 25 L 75 25 L 75 75 L 25 75 Z')!;
		const evenOdd = blurredSilhouette([{ segments: ring, fill: 'evenodd' }], 1, black)!;
		expect(alphaAt(evenOdd, 50, 50)).toBeLessThan(0.05);
		expect(alphaAt(evenOdd, 10, 50)).toBeGreaterThan(0.95);
		const nonzero = blurredSilhouette([{ segments: ring, fill: 'nonzero' }], 1, black)!;
		expect(alphaAt(nonzero, 50, 50)).toBeGreaterThan(0.95);
		const line = parsePath('M 0 50 L 100 50')!;
		const stroked = blurredSilhouette([{ segments: line, stroke: 20 }], 1, black)!;
		expect(alphaAt(stroked, 50, 50)).toBeGreaterThan(0.95);
		expect(alphaAt(stroked, 50, 62.5)).toBeLessThan(0.05);
		const moved = blurredSilhouette(
			[{ segments: square, fill: 'nonzero', matrix: [2, 0, 0, 2, 50, 0] }],
			1,
			black,
		)!;
		expect(alphaAt(moved, 200, 100)).toBeGreaterThan(0.95);
		expect(moved.x).toBeCloseTo(47);
		expect(alphaAt(moved, 47.5, 100)).toBeLessThan(0.05);
	});

	it('colours every sample with the shadow colour and returns nothing for no area', () => {
		const mask = blurredSilhouette([{ segments: square, fill: 'nonzero' }], 4, {
			r: 1,
			g: 0.5,
			b: 0,
		})!;
		expect([...mask.resource.data.subarray(0, 6)]).toEqual([255, 128, 0, 255, 128, 0]);
		expect(blurredSilhouette([{ segments: square }], 4, black)).toBeUndefined();
		expect(blurredSilhouette([], 4, black)).toBeUndefined();
	});

	it('flattens curves into short pieces', () => {
		const curve = parsePath('M 0 0 C 0 100 100 100 100 0')!;
		const [line] = flattenSegments(curve, undefined, 1);
		expect(line!.points.length).toBeGreaterThan(20);
		expect(line!.points.at(-1)).toEqual([100, 0]);
	});
});

describe('drawBlurredSilhouette', () => {
	it('paints one image with a soft mask at the shadow opacity', () => {
		const page = new PdfVectorPage(200, 200);
		expect(
			drawBlurredSilhouette(page, [{ segments: square, fill: 'nonzero' }], 6, black, 0.5),
		).toBe(true);
		expect(page.content).toMatch(/^q\n\/GS0 gs\n136 0 0 -136 -18 118 cm\n\/Im0 Do\nQ$/);
		const raw = Buffer.from(buildVectorPdf([page])).toString('latin1');
		expect(raw).toMatch(/\/GS0 << \/ca 0\.5 \/CA 0\.5 >>/);
		const smask =
			/\/Subtype \/Image [^>]*\/ColorSpace \/DeviceGray[^>]*\/Length (\d+) >>\nstream\n/.exec(raw);
		expect(smask).not.toBeNull();
		expect(raw).toMatch(/\/SMask \d+ 0 R/);
		const start = smask!.index + smask![0].length;
		const alpha = inflateSync(Buffer.from(raw.slice(start, start + Number(smask![1])), 'latin1'));
		expect(Math.max(...alpha)).toBe(255);
		expect(alpha[0]).toBe(0);
	});
});
