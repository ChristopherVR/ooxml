// @vitest-environment happy-dom
import type { ImagePptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { ElementRenderContext } from '../types';
import { renderImageElement } from './image';

function context(): ElementRenderContext {
	return {
		document,
		mediaDataUrls: new Map(),
		renderElement: () => null,
	} as unknown as ElementRenderContext;
}

function picture(shapeStyle?: ImagePptxElement['shapeStyle']): ImagePptxElement {
	return {
		type: 'image',
		id: 'pic-frame',
		x: 0,
		y: 0,
		width: 200,
		height: 150,
		imageData: 'data:image/png;base64,AA==',
		...(shapeStyle ? { shapeStyle } : {}),
	};
}

const FRAME = { strokeWidth: 10.67, strokeFillMode: 'solid', strokeColor: '#FF0000' } as const;
const SHADOW = {
	shadowColor: '#000000',
	shadowOpacity: 0.6,
	shadowBlur: 5.33,
	shadowAngle: 45,
	shadowDistance: 10.67,
} as const;

describe('renderImageElement picture frame (a:ln + a:outerShdw)', () => {
	it('paints the frame as an outline overlay outside the clipped surface', () => {
		const el = renderImageElement(picture(FRAME), 0, context()) as HTMLElement;
		expect(el.style.overflow).toBe('visible');
		const surface = el.querySelector<HTMLElement>(':scope > .pptxv-image-surface');
		expect(surface?.style.overflow).toBe('hidden');
		expect(surface?.querySelector('img')).not.toBeNull();
		const path = el.querySelector(':scope > svg path');
		expect(path?.getAttribute('stroke')).toBe('#FF0000');
		expect(Number(path?.getAttribute('stroke-width'))).toBeCloseTo(10.67);
	});

	it('casts the outer shadow from the unclipped box', () => {
		const el = renderImageElement(picture({ ...FRAME, ...SHADOW }), 0, context()) as HTMLElement;
		expect(el.style.filter).toMatch(/drop-shadow\(/);
		expect(el.querySelector<HTMLElement>('img')?.style.filter ?? '').not.toMatch(/drop-shadow/);
	});

	it("hangs a framed picture's reflection off the box, not the clipped surface", () => {
		const el = renderImageElement(
			picture({ ...FRAME, reflectionStartOpacity: 0.5, reflectionDistance: 4 }),
			0,
			context(),
		) as HTMLElement;
		expect(el.querySelector(':scope > .pptxv-reflection')).not.toBeNull();
		expect(el.querySelector('.pptxv-image-surface .pptxv-reflection')).toBeNull();
	});

	it('keeps the single clipped box for a picture without a frame', () => {
		const el = renderImageElement(picture(), 0, context()) as HTMLElement;
		expect(el.style.overflow).toBe('hidden');
		expect(el.querySelector('.pptxv-image-surface')).toBeNull();
		expect(el.querySelector(':scope > img')).not.toBeNull();
	});
});
