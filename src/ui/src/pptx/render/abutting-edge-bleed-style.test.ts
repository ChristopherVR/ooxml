import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { abuttingEdgeBleedStyle } from './abutting-edge-bleed-style';

const element = { id: 'o8', type: 'shape', x: 581, y: 0, width: 350, height: 720 } as PptxElement;
const gradient = 'linear-gradient(90deg, rgb(0, 64, 94) 0%, rgba(51, 102, 126, 0) 95%)';

describe('abuttingEdgeBleedStyle', () => {
	it('is empty without a bleed', () => {
		expect(abuttingEdgeBleedStyle(element, undefined, { backgroundColor: '#000' })).toEqual({});
		expect(abuttingEdgeBleedStyle(element, {}, { backgroundColor: '#000' })).toEqual({});
	});

	it('grows a background-colour box on its grown side only', () => {
		expect(
			abuttingEdgeBleedStyle(element, { left: { color: '#000' } }, { backgroundColor: '#000' }),
		).toEqual({
			left: '579px',
			top: '0px',
			width: '352px',
			height: '720px',
		});
	});

	it('keeps a gradient at its authored box and lays the edge colour under the strip', () => {
		const style = abuttingEdgeBleedStyle(
			element,
			{ left: { color: '#00405E' } },
			{ backgroundImage: gradient },
		);
		expect(style).toEqual({
			left: '579px',
			top: '0px',
			width: '352px',
			height: '720px',
			backgroundImage: `${gradient}, linear-gradient(#00405E, #00405E)`,
			backgroundSize: '350px 720px, 3px 720px',
			backgroundPosition: '2px 0px, 0px 0px',
			backgroundRepeat: 'no-repeat',
		});
	});

	it('carries a gradient that runs along the edge into the strip with the image itself', () => {
		const style = abuttingEdgeBleedStyle(element, { bottom: {} }, { backgroundImage: gradient });
		expect(style['backgroundImage']).toBe(`${gradient}, ${gradient}`);
		expect(style['backgroundSize']).toBe('350px 720px, 350px 3px');
		expect(style['backgroundPosition']).toBe('0px 0px, 0px 719px');
		expect(style['height']).toBe('722px');
	});

	it('reads and writes a kebab-case style map in its own spelling', () => {
		const style = abuttingEdgeBleedStyle(
			element,
			{ left: { color: '#00405E' } },
			{ 'background-image': gradient },
		);
		expect(style['background-image']).toBe(`${gradient}, linear-gradient(#00405E, #00405E)`);
		expect(style['background-size']).toBe('350px 720px, 3px 720px');
		expect(style['backgroundImage']).toBeUndefined();
		expect(
			abuttingEdgeBleedStyle(
				element,
				{ left: {} },
				{ 'background-color': '#000', 'box-shadow': '0 0 1px #000' },
			),
		).toEqual({});
	});

	it('accepts the border-box origin a binding writes for a borderless box', () => {
		const style = abuttingEdgeBleedStyle(
			element,
			{ left: { color: '#000' } },
			{
				backgroundColor: '#000',
				backgroundOrigin: 'border-box',
				borderWidth: 0,
			},
		);
		expect(style['width']).toBe('352px');
		expect(
			abuttingEdgeBleedStyle(
				element,
				{ left: { color: '#000' } },
				{ backgroundColor: '#000', backgroundOrigin: 'content-box' },
			),
		).toEqual({});
		expect(
			abuttingEdgeBleedStyle(
				element,
				{ left: { color: '#000' } },
				{ backgroundImage: gradient, padding: '4px' },
			),
		).toEqual({});
	});

	it('stays out of a box that paints anything besides a plain background fill', () => {
		const bleed = { left: { color: '#000' } };
		for (const painted of [
			{},
			{ backgroundColor: 'transparent', backgroundImage: 'none' },
			{ backgroundColor: '#000', boxShadow: '0 0 4px #000' },
			{ backgroundColor: '#000', borderWidth: '2px' },
			{ backgroundImage: gradient, backgroundSize: '10px 10px' },
			{ backgroundImage: gradient, backgroundRepeat: 'repeat' },
			{ backgroundColor: '#000', clipPath: 'inset(0)' },
			{ backgroundColor: '#000', filter: 'blur(1px)' },
		]) {
			expect(abuttingEdgeBleedStyle(element, bleed, painted), JSON.stringify(painted)).toEqual({});
		}
	});
});
