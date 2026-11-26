import { expect, it } from 'vitest';

import { ShapeBuilder } from '../../builders/sdk';
import { PptxHandler } from '../../PptxHandler';

async function makeDeck(): Promise<ArrayBuffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 3 });
	try {
		for (const [index, slide] of data.slides.entries()) {
			slide.elements = [
				ShapeBuilder.create('rect')
					.text(`Slide ${index + 1}`)
					.build(),
			];
			slide.isDirty = true;
		}
		return Uint8Array.from(await handler.save(data.slides)).buffer;
	} finally {
		handler.dispose();
	}
}

it('publishes a slide before a later slide hits the shape limit', async () => {
	const handler = new PptxHandler();
	const events: number[] = [];
	try {
		await expect(
			handler.load(await makeDeck(), {
				limits: { maxShapes: 1 },
				async onSlideLoaded({ index, total, slide }) {
					expect(total).toBe(3);
					expect(slide.elements[0]).toMatchObject({ text: 'Slide 1' });
					await Promise.resolve();
					events.push(index);
				},
			}),
		).rejects.toMatchObject({ code: 'PPTX_LOAD_LIMIT', limit: 'maxShapes' });
		expect(events).toEqual([0]);
	} finally {
		handler.dispose();
	}
});

it('returns the complete deck for editing and preserves the deferred slides on save', async () => {
	const handler = new PptxHandler();
	const reopened = new PptxHandler();
	try {
		const order: string[] = [];
		const data = await handler.load(await makeDeck(), {
			onSlideLoaded({ slide }) {
				order.push(slide.id);
			},
		});
		expect(data.slides.map(({ id }) => id)).toEqual(order);
		const first = data.slides[0];
		if (!first) throw new Error('The test deck needs a first slide');
		first.elements = [ShapeBuilder.create('rect').text('Edited first slide').build()];
		first.isDirty = true;
		const bytes = Uint8Array.from(await handler.save(data.slides)).buffer;
		const saved = await reopened.load(bytes);
		expect(saved.slides.map((slide) => slide.elements[0])).toMatchObject([
			{ text: 'Edited first slide' },
			{ text: 'Slide 2' },
			{ text: 'Slide 3' },
		]);
	} finally {
		handler.dispose();
		reopened.dispose();
	}
});
