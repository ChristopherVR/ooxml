// @vitest-environment jsdom
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createTranslator } from './i18n';
import { createDefaultRegistry, renderSlideStage } from './render';

/**
 * A picture's own frame (`a:ln`) and outer shadow, painted through the slide
 * stage this binding renders. Before the shared picture-frame fix the vanilla
 * viewer drew the bare bitmap: no outline and no shadow, so every Picture
 * Styles pick with a frame changed the saved file and nothing on screen.
 */
function stageWith(picture: PptxElement): HTMLElement {
	return renderSlideStage({
		document,
		slide: { id: 's1', rId: 'r1', slideNumber: 1, elements: [picture] } as PptxSlide,
		canvasSize: { width: 960, height: 540 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator('en'),
		scale: 1,
		interactive: false,
	});
}

const PICTURE = {
	type: 'picture',
	id: 'pic',
	x: 100,
	y: 75,
	width: 200,
	height: 150,
	imageData: 'data:image/png;base64,AA==',
} as PptxElement;

describe('vanilla: a picture paints its own frame and outer shadow', () => {
	it('strokes the frame outside the clipped surface and casts the shadow', () => {
		const stage = stageWith({
			...PICTURE,
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
		} as PptxElement);
		const box = stage.querySelector<HTMLElement>('[data-element-id="pic"]');
		expect(box?.style.overflow).toBe('visible');
		expect(box?.style.filter).toMatch(/drop-shadow\(/);
		expect(box?.querySelector('.pptxv-image-surface img')).not.toBeNull();
		const path = box?.querySelector(':scope > svg path');
		expect(path?.getAttribute('stroke')).toBe('#FF0000');
	});

	it('keeps an unframed picture in its single clipped box', () => {
		const box = stageWith(PICTURE).querySelector<HTMLElement>('[data-element-id="pic"]');
		expect(box?.style.overflow).toBe('hidden');
		expect(box?.querySelector('.pptxv-image-surface')).toBeNull();
	});
});
