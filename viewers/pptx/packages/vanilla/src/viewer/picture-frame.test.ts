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

it('renders an unfilled freeform with its arrowhead and no clipping box', () => {
	const element = {
		id: 'freeform',
		type: 'shape',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		pathData: 'M 0 100 L 100 0',
		pathWidth: 100,
		pathHeight: 100,
		shapeStyle: {
			fillMode: 'none',
			strokeWidth: 3,
			strokeColor: '#1F4E79',
			connectorEndArrow: 'triangle',
		},
	} as PptxElement;
	const box = stageWith(element).querySelector<HTMLElement>('[data-element-id="freeform"]')!;
	expect(box.querySelector('path[marker-end]')).not.toBeNull();
	expect(box.querySelector('marker')?.getAttribute('markerWidth')).toBe('3');
	expect(box.style.clipPath).toBe('');
});

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
		expect(
			box!.querySelector('img')!.compareDocumentPosition(path!) & Node.DOCUMENT_POSITION_FOLLOWING,
		).toBeTruthy();
	});

	it('keeps an unframed picture in its single clipped box', () => {
		const box = stageWith(PICTURE).querySelector<HTMLElement>('[data-element-id="pic"]');
		expect(box?.style.overflow).toBe('hidden');
		expect(box?.querySelector('.pptxv-image-surface')).toBeNull();
	});
});

describe('vanilla: a shape paints the shadow and glow profiles of PowerPoint', () => {
	// Measured on PowerPoint 16 (see src/ui/src/pptx/render/effect-profile.ts): the
	// blurred shadow is a Gaussian of 0.326 blurRad (CSS length 0.654 blurRad), the
	// glow one box-shadow layer grown by 0.417 R and blurred by R / 3.
	it('paints a blurRad 30 shadow at 20px and a rad 30 glow as one spread layer', () => {
		const box = stageWith({
			id: 'fx',
			type: 'shape',
			shapeType: 'rect',
			x: 100,
			y: 75,
			width: 200,
			height: 150,
			shapeStyle: {
				fillColor: '#C07020',
				shadowColor: '#000000',
				shadowBlur: 30,
				shadowOpacity: 1,
				outerShadowXml: {},
				glowColor: '#FF0000',
				glowRadius: 30,
				glowOpacity: 0.6,
			},
		} as PptxElement).querySelector<HTMLElement>('[data-element-id="fx"]');
		expect(box?.style.boxShadow).toBe(
			'0px 0px 20px rgba(0, 0, 0, 1), 0 0 10px 12.5px rgba(255, 0, 0, 0.6)',
		);
		expect(box?.style.filter ?? '').not.toContain('drop-shadow');
	});
});
