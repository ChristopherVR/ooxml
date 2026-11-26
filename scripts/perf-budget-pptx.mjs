/** Generated, editable decks shared by the work budget and the browser benchmark. */
export const PPTX_OPEN_SMALL = 4;
export const PPTX_OPEN_LARGE = 180;
export const PPTX_SHAPES_PER_SLIDE = 12;
export const PPTX_OPEN_WORK = 'work/pptx/large deck missed first-slide previews';

export async function makePptxOpeningDeck(slideCount) {
	const { PptxHandler, ShapeBuilder } = await import('../src/core/dist/pptx/index.mjs');
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slideCount });
	try {
		for (const [index, slide] of data.slides.entries()) {
			slide.elements = Array.from({ length: PPTX_SHAPES_PER_SLIDE }, (_, shape) =>
				ShapeBuilder.create('rect')
					.position(40 + (shape % 3) * 305, 56 + Math.floor(shape / 3) * 112)
					.size(265, 77)
					.text(`Deck slide ${index + 1} shape ${shape + 1}`)
					.build(),
			);
			slide.isDirty = true;
		}
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

export async function measurePptxOpeningWork() {
	const { getPptxLoadLimitViolation, PptxHandler } =
		await import('../src/core/dist/pptx/index.mjs');
	const bytes = await makePptxOpeningDeck(PPTX_OPEN_LARGE);
	const handler = new PptxHandler();
	let firstPreview = false;
	try {
		await handler.load(Uint8Array.from(bytes).buffer, {
			// Shapes on the first slide fit. The next slide must hit this guard
			// after the preview callback, which proves parsing really suspended.
			limits: { maxShapes: PPTX_SHAPES_PER_SLIDE },
			onSlideLoaded({ index, total, slide }) {
				firstPreview =
					index === 0 &&
					total === PPTX_OPEN_LARGE &&
					slide.elements.length === PPTX_SHAPES_PER_SLIDE;
			},
		});
	} catch (error) {
		if (getPptxLoadLimitViolation(error)?.limit === 'maxShapes') {
			return firstPreview ? 0 : 1;
		}
		throw error;
	} finally {
		handler.dispose();
	}
	throw new Error('The large PowerPoint fixture never reached its second slide shape limit');
}
