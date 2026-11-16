import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { getComputedFillStyle } from './fill-style';
import { shapeFillModeChange } from './inspector/shape-fill-fields';
import { getSlideBackgroundStyle } from './slide-background';

const slide = (extra: Partial<PptxSlide> = {}): PptxSlide =>
	({ id: 's1', rId: 'r1', slideNumber: 1, elements: [], ...extra }) as PptxSlide;

function shape(style: Record<string, unknown>): PptxElement {
	return {
		type: 'shape',
		id: 'sh1',
		x: 40,
		y: 25,
		width: 100,
		height: 50,
		shapeType: 'rect',
		shapeStyle: style,
	} as PptxElement;
}

describe('slide background fill drawn live', () => {
	it('publishes the slide background as custom properties on the stage', () => {
		const solid = getSlideBackgroundStyle(slide({ backgroundColor: '#112233' }));
		expect(solid['--pptx-slide-bg-color']).toBe('#112233');
		expect(solid['--pptx-slide-bg-image']).toBe('none');
		const gradient = 'linear-gradient(90deg, #000 0%, #fff 100%)';
		const sized = getSlideBackgroundStyle(slide({ backgroundGradient: gradient }), {
			widthPx: 960,
			heightPx: 540,
		});
		expect(sized['--pptx-slide-bg-image']).toBe(gradient);
		expect(sized['--pptx-slide-bg-size']).toBe('960px 540px');
		expect(sized['--pptx-slide-bg-repeat']).toBe('no-repeat');
	});

	it('paints a shape switched to the slide background with those properties', () => {
		const style = shapeFillModeChange({ fillColor: '#ff0000', fillMode: 'solid' }, 'background');
		const fill = getComputedFillStyle(shape(style));
		expect(fill).toStrictEqual({
			backgroundColor: 'var(--pptx-slide-bg-color, transparent)',
			backgroundImage: 'var(--pptx-slide-bg-image, none)',
			backgroundSize: 'var(--pptx-slide-bg-size, 100% 100%)',
			backgroundRepeat: 'var(--pptx-slide-bg-repeat, no-repeat)',
			backgroundPosition: '-40px -25px',
		});
	});

	it('keeps the copy a loaded background-filled shape already carries', () => {
		const fill = getComputedFillStyle(
			shape({ useBackgroundFill: true, fillMode: 'solid', fillColor: '#00ff00' }),
		);
		expect(fill?.backgroundColor).not.toContain('var(');
	});
});
