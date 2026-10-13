import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { getImageSurfaceStyle } from '../../utils/image-style';
import { getShapeVisualStyle } from '../../utils/shape-visual-style';
import { shapeParams } from './element-shape-params';
import { ShapeEffectOverlay } from './ShapeEffectOverlay';

/**
 * React is the reference the other bindings now follow for a picture's own
 * frame (`a:ln`) and outer shadow (shared `getPictureFrameLayers`): the
 * element container is unclipped and carries the composite shadow, the
 * picture's pixels sit in a clipped surface, and `ShapeEffectOverlay` strokes
 * the centred frame beside that surface. This pins the three pieces.
 */
const PICTURE = {
	type: 'picture',
	id: 'pic-frame',
	x: 0,
	y: 0,
	width: 200,
	height: 150,
	imageData: 'data:image/png;base64,AA==',
	shapeStyle: {
		strokeWidth: 10.67,
		strokeFillMode: 'solid',
		strokeColor: '#FF0000',
		shadowColor: '#000000',
		shadowOpacity: 0.6,
		shadowBlur: 5.33,
		shadowAngle: 45,
		shadowDistance: 10.67,
	},
} as unknown as PptxElement;

describe('react picture frame (a:ln + a:outerShdw)', () => {
	it('casts the shadow from the container and paints no CSS border there', () => {
		const { hf, fc, sw, sc } = shapeParams(PICTURE);
		const style = getShapeVisualStyle(PICTURE, hf, fc, sw, sc);
		expect(String(style.filter)).toMatch(/drop-shadow\(/);
		expect(style.borderWidth).toBe(0);
		expect(style.overflow).not.toBe('hidden');
	});

	it('clips the pixels on the surface, not the container', () => {
		expect(getImageSurfaceStyle(PICTURE).overflow).toBe('hidden');
	});

	it('strokes the centred frame in the effect overlay', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={PICTURE} />);
		expect(html).toMatch(/<path[^>]*stroke="#FF0000"/);
		expect(html).toMatch(/stroke-width="10\.67"/);
	});
});
