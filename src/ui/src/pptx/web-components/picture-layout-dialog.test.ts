import { describe, expect, it } from 'vitest';
import type { PptxElement, PptxImageLikeElement, PptxSlide } from 'ooxml-core/pptx';

import { pictureLayoutTargets, replacePicturesWithDiagram } from './picture-layout-dialog';

const picture = (id: string, extra: Partial<PptxImageLikeElement> = {}): PptxImageLikeElement =>
	({ id, type: 'picture', x: 0, y: 0, width: 10, height: 10, ...extra }) as PptxImageLikeElement;
const shape = { id: 'shape', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;

describe('Picture Layout targets', () => {
	const slide: PptxSlide = {
		id: 's1',
		rId: 'rId2',
		slideNumber: 1,
		elements: [picture('a'), shape, picture('b'), picture('c')],
	};

	it('converts every selected picture on the slide, never other elements', () => {
		const a = slide.elements[0] as PptxImageLikeElement;
		expect(pictureLayoutTargets(slide, a, ['a', 'shape', 'c']).map(({ id }) => id)).toEqual([
			'a',
			'c',
		]);
		expect(pictureLayoutTargets(slide, a, undefined).map(({ id }) => id)).toEqual(['a']);
	});

	it('puts the diagram where the topmost converted picture was', () => {
		const diagram = { id: 'diagram', type: 'smartArt', x: 0, y: 0, width: 1, height: 1 };
		const targets = [slide.elements[0], slide.elements[2]] as PptxImageLikeElement[];
		expect(
			replacePicturesWithDiagram(slide.elements, targets, diagram as PptxElement).map(
				({ id }) => id,
			),
		).toEqual(['shape', 'diagram', 'c']);
	});
});
