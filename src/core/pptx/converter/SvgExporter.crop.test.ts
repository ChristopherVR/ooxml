import { readFile } from 'node:fs/promises';

import { describe, expect, it } from 'vitest';

import { handleExportSvg } from '../cli/commands';
import type { ImagePptxElement, PicturePptxElement } from '../core/types/elements';
import { SvgExporter } from './SvgExporter';

const imageData = 'data:image/png;base64,iVBOR';

function exportImage(overrides: Partial<ImagePptxElement | PicturePptxElement> = {}): string {
	return SvgExporter.exportSlide(
		{
			id: 'slide1',
			rId: 'rId1',
			slideNumber: 1,
			elements: [
				{
					type: 'image',
					id: 'picture1',
					x: 10,
					y: 10,
					width: 200,
					height: 100,
					imageData,
					...overrides,
				},
			],
		},
		220,
		120,
	);
}

describe('SVG picture source crop', () => {
	it.each(['image', 'picture'] as const)('clips the source for a %s element', (type) => {
		const svg = exportImage({ type, cropRight: 0.5 });
		expect(svg).toContain('width="200" height="100" viewBox="0 0 100 100"');
		expect(svg).toContain('preserveAspectRatio="none" overflow="hidden"><image');
	});

	it('applies both axes while retaining the destination frame', () => {
		const svg = exportImage({ cropLeft: 0.25, cropTop: 0.1, cropRight: 0.25, cropBottom: 0.2 });
		expect(svg).toContain('viewBox="50 10 100 70"');
		expect(svg).toContain('overflow="hidden"');
	});

	it('preserves padding from negative source insets', () => {
		const svg = exportImage({ cropLeft: -0.5, cropRight: -0.5 });
		expect(svg).toContain('viewBox="-100 0 400 100"');
	});

	it('leaves uncropped output unchanged', () => {
		expect(exportImage({ cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0 })).toBe(
			exportImage(),
		);
	});

	it.each([
		{ cropLeft: Number.NaN },
		{ cropTop: Number.POSITIVE_INFINITY },
		{ cropRight: 1 },
		{ cropTop: 0.6, cropBottom: 0.6 },
	])('falls back to the image for invalid crop geometry %j', (crop) => {
		expect(exportImage(crop)).toBe(exportImage());
	});

	it('retains rotation, flips and image effects around a crop', () => {
		const overrides = { rotation: 30, flipHorizontal: true, imageEffects: { grayscale: true } };
		const plain = exportImage(overrides);
		const cropped = exportImage({ ...overrides, cropRight: 0.5 });
		expect(cropped).toContain('viewBox="0 0 100 100"');
		expect(cropped.match(/<g transform="[^"]*"/g)).toStrictEqual(
			plain.match(/<g transform="[^"]*"/g),
		);
		expect(cropped.match(/<filter[^>]*>.*?<\/filter>/g)).toStrictEqual(
			plain.match(/<filter[^>]*>.*?<\/filter>/g),
		);
		expect(cropped).toContain('filter="url(#');
	});

	it('keeps the crop when the public CLI exports a saved PPTX', async () => {
		const bytes = await readFile(
			new URL('../__tests__/fixtures/svg-picture-crop.pptx', import.meta.url),
		);
		const { svgs } = await handleExportSvg(bytes);
		expect(svgs).toHaveLength(1);
		expect(svgs[0]).toContain('viewBox="0 0 330 396"');
		expect(svgs[0]).toContain('data:image/png;base64,');
	});
});
