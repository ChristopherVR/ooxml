import { expect, it } from 'vitest';
import type { PptxElement, PptxSlide } from '../../index';
import { collectElementText, searchSlides } from './slide-search';

it('collects visible rich text once, without splitting words at formatting runs', () => {
	const element: PptxElement = {
		id: 'text',
		type: 'text',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		text: 'Annual report',
		textSegments: [
			{ text: 'Annual re', style: { bold: true } },
			{ text: 'port', style: { italic: true } },
		],
	};
	expect(collectElementText(element)).toBe('Annual report');
	const slides: PptxSlide[] = [{ id: 's', rId: 'r', slideNumber: 1, elements: [element] }];
	expect(searchSlides(slides, 'report')[0]?.matchCount).toBe(1);
});
