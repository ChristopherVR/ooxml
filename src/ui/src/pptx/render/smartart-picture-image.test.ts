// @vitest-environment happy-dom
import { afterEach, expect, it, vi } from 'vitest';
import { smartArtPictureImage, refreshSmartArtPictureImage } from './smartart-picture-image';
import * as pixels from '../../picture-pixels';
import type { PptxSmartArtDrawingShape } from 'ooxml-core/pptx';
afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

it('shares cropped, effect-bearing image paint clipped to the authored preset', () => {
	const result = smartArtPictureImage({
		id: 'picture',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'ellipse',
		fillImageUrl: 'data:image/png;base64,AA==',
		pictureFill: {
			kind: 'picture',
			crop: { cropLeft: 0.125, cropRight: 0.25 },
			effects: [{ kind: 'grayscale' }],
		},
	});
	expect(result).toMatch(/^data:image\/svg\+xml,/);
	const svg = new DOMParser().parseFromString(
		decodeURIComponent(result!.split(',').slice(1).join(',')),
		'image/svg+xml',
	);
	expect(svg.querySelector('pattern image')?.getAttribute('width')).toBe('1.6');
	expect(svg.querySelector('feColorMatrix')).not.toBeNull();
	expect(svg.documentElement.lastElementChild?.localName).toBe('path');
	expect(svg.documentElement.lastElementChild?.getAttribute('fill')).toMatch(
		/^url\(#office-picture-/,
	);
});

it.each(['mounted', 'removed', 'subtree', 'changed', 'replaced'] as const)(
	'invalidates only a current %s image after asynchronous pixel paint',
	async (state) => {
		let complete!: (value: pixels.PicturePixelImage) => void;
		const probe = vi.spyOn(pixels, 'probePicturePixels').mockReturnValue(
			new Promise((resolve) => {
				complete = resolve;
			}),
		);
		const cached = vi.spyOn(pixels, 'cachedPicturePixels').mockReturnValue(undefined);
		const source: PptxSmartArtDrawingShape = {
			id: 'async',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			fillImageUrl: 'data:image/png;base64,AA==',
			pictureFill: {
				kind: 'picture',
				effects: [{ kind: 'standard', name: 'hsl', values: { saturation: 1, luminance: 0 } }],
			},
		};
		const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
		const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
		image.setAttribute('href', 'pending');
		svg.append(image);
		document.body.append(svg);
		refreshSmartArtPictureImage({ currentTarget: image }, source);
		if (state === 'removed') image.remove();
		if (state === 'subtree') svg.remove();
		if (state === 'replaced') image.setAttribute('href', 'replacement');
		if (state === 'changed') source.fillImageUrl = 'data:image/png;base64,NEW';
		const result = { url: 'data:image/png;base64,PROCESSED', width: 48, height: 32 };
		cached.mockReturnValue(result);
		complete(result);
		await Promise.resolve();
		if (state === 'mounted') {
			expect(image.getAttribute('href')).toMatch(/^data:image\/svg\+xml,/);
			expect(image.getAttribute('data-office-picture-pixels-ready')).toBe('true');
			refreshSmartArtPictureImage({ currentTarget: image }, source);
			expect(probe).toHaveBeenCalledTimes(1);
		} else
			expect(image.getAttribute('href')).toBe(state === 'replaced' ? 'replacement' : 'pending');
	},
);

it('serializes native-size tiles synchronously without waiting for an image probe', () => {
	const result = smartArtPictureImage({
		id: 'tile',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		fillImageUrl: 'data:image/png;base64,AA==',
		fillImageWidth: 48,
		fillImageHeight: 32,
		pictureFill: { kind: 'picture', layout: { tile: { scaleX: 0.5, scaleY: 1 } } },
	});
	const svg = new DOMParser().parseFromString(
		decodeURIComponent(result!.split(',').slice(1).join(',')),
		'image/svg+xml',
	);
	expect(svg.querySelector('pattern')?.getAttribute('width')).toBe('24');
	expect(svg.querySelector('pattern')?.getAttribute('height')).toBe('32');
	expect(svg.querySelector('pattern image')).not.toBeNull();
});
