import { PresentationBuilder } from 'pptx-viewer-core';

import { patchOpacitySetPackage } from '../../../src/core/pptx/__tests__/helpers/opacity-set-deck';

/** Published SDK entry avoids importing unbundled core JSON in Node browser workers. */
export async function opacitySetDeck(hold: boolean): Promise<Buffer> {
	const { handler, createSlide } = await PresentationBuilder.create();
	try {
		const slide = createSlide('Blank')
			.addShape('rect', {
				x: 100,
				y: 100,
				width: 400,
				height: 200,
				fill: { type: 'solid', color: '#00AA00' },
			})
			.build();
		const element = slide.elements[0]!;
		if (element.type !== 'shape') throw new Error('Expected shape');
		element.text = 'Persistent text';
		element.textSegments = [{ text: element.text, style: { color: '#000000', fontSize: 24 } }];
		const next = createSlide('Blank')
			.addText('Next slide', { x: 100, y: 100, width: 400, height: 100 })
			.build();
		return patchOpacitySetPackage(await handler.save([slide, next]), hold, hold);
	} finally {
		handler.dispose();
	}
}
