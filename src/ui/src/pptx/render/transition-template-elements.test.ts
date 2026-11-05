import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { transitionTemplateElements } from './transition-template-elements';

const background = (id: string): PptxElement =>
	({ id, type: 'picture', x: 0, y: 0, width: 1280, height: 720 }) as PptxElement;

const slide = (id: string, showMasterShapes?: boolean): PptxSlide =>
	({
		id,
		elements: [],
		...(showMasterShapes === undefined ? {} : { showMasterShapes }),
	}) as PptxSlide;

// Issue #141, slides 1 -> 2: the slides sit on different layouts, each with
// its own background picture. The leaving slide has to keep its own.
describe('transitionTemplateElements', () => {
	const byId = {
		title: [background('title-layout-bg')],
		content: [background('content-layout-bg')],
	};

	it('puts each layer behind its own slide layout', () => {
		const layers = transitionTemplateElements(slide('title'), slide('content'), byId);
		expect(layers.outgoing.map((element) => element.id)).toEqual(['title-layout-bg']);
		expect(layers.incoming.map((element) => element.id)).toEqual(['content-layout-bg']);
	});

	it('honours Hide Background Graphics on each slide separately', () => {
		const layers = transitionTemplateElements(slide('title', false), slide('content'), byId);
		expect(layers.outgoing).toEqual([]);
		expect(layers.incoming.map((element) => element.id)).toEqual(['content-layout-bg']);
	});

	it('paints no template for a slide without one, or without an arriving slide', () => {
		expect(transitionTemplateElements(slide('blank'), undefined, byId)).toEqual({
			outgoing: [],
			incoming: [],
		});
	});
});
