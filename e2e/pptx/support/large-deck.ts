import { PptxHandler } from 'pptx-viewer-core';

/** Slides in {@link largeDeck}: above the rail's 50-slide virtualization threshold. */
export const LARGE_DECK_SLIDES = 80;

/** A deck of {@link LARGE_DECK_SLIDES} slides with one titled text box each. */
export async function largeDeck(): Promise<Buffer> {
	const { handler, createSlide } = await PptxHandler.create({ initialSlideCount: 0 });
	try {
		const slides = Array.from({ length: LARGE_DECK_SLIDES }, (_, index) =>
			createSlide('Blank')
				.addText(`Large deck slide ${index + 1}`, { x: 60, y: 60, width: 480, height: 60 })
				.build(),
		);
		return Buffer.from(await handler.save(slides));
	} finally {
		handler.dispose();
	}
}
