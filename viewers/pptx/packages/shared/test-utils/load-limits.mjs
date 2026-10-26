/** Slides in {@link loadLimitDeck}. */
export const LOAD_LIMIT_DECK_SLIDES = 3;

/** The English load-error text each binding shows for {@link loadLimitDeck} under `maxSlides: 2`. */
export const OVER_SLIDE_LIMIT_MESSAGE =
	'This presentation has 3 slides, more than the 2 this viewer accepts.';

/** A real three-slide archive, for the host `limits` option tests of every binding. */
export async function loadLimitDeck(Handler) {
	const { handler, createSlide } = await Handler.create({ initialSlideCount: 0 });
	try {
		const slides = [];
		for (let index = 0; index < LOAD_LIMIT_DECK_SLIDES; index++) {
			slides.push(
				createSlide('Blank')
					.addText(`Slide ${index + 1}`, { x: 50, y: 50, width: 400, height: 60 })
					.build(),
			);
		}
		return await handler.save(slides);
	} finally {
		handler.dispose();
	}
}
