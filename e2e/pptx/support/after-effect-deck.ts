import { readFile } from 'node:fs/promises';

import { PptxHandler, PresentationBuilder } from 'pptx-viewer-core';

import { fixture } from './deck';

/** Retain the four native effects, with one real following click for after-state boundaries. */
export async function afterEffectDeck(): Promise<Buffer> {
	const handler = new PptxHandler();
	const builder = await PresentationBuilder.create();
	try {
		const source = await readFile(fixture('animation-after-effect.pptx'));
		const data = await handler.load(
			source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
		);
		const slide = data.slides[0]!;
		const dummy = builder
			.createSlide('Blank')
			.addShape('rect', {
				x: 650,
				y: 400,
				width: 120,
				height: 50,
				fill: { type: 'solid', color: '#FFFFFF' },
			})
			.build().elements[0]!;
		if (dummy.type !== 'shape') throw new Error('Expected a shape');
		dummy.id = `${slide.id}-shape-${slide.elements.length}`;
		dummy.text = 'Next trigger';
		dummy.textSegments = [{ text: dummy.text, style: { color: '#000000', fontSize: 12 } }];
		slide.elements.push(dummy);
		slide.animations = [
			...(slide.animations ?? []),
			{ elementId: dummy.id, entrance: 'appear', durationMs: 1, trigger: 'onClick', order: 4 },
		];
		return Buffer.from(await handler.save(data.slides));
	} finally {
		handler.dispose();
		builder.handler.dispose();
	}
}
