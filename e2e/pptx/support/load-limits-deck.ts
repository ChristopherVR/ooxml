import { PptxHandler } from 'pptx-viewer-core';

/** Slides in {@link loadLimitsDeck}. */
export const LOAD_LIMITS_DECK_SLIDES = 3;

/** A deck of {@link LOAD_LIMITS_DECK_SLIDES} slides with one titled text box each. */
export async function loadLimitsDeck(): Promise<Buffer> {
	const { handler, createSlide } = await PptxHandler.create({ initialSlideCount: 0 });
	try {
		const slides = Array.from({ length: LOAD_LIMITS_DECK_SLIDES }, (_, index) =>
			createSlide('Blank')
				.addText(`Load limits slide ${index + 1}`, { x: 60, y: 60, width: 480, height: 60 })
				.build(),
		);
		return Buffer.from(await handler.save(slides));
	} finally {
		handler.dispose();
	}
}
