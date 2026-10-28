// @vitest-environment happy-dom
import { expect, it } from 'vitest';
import { smartArtPictureImage } from './smartart-picture-image';

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
