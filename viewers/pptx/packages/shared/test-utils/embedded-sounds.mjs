/** A silent 8 kHz WAV, embedded as each sound of {@link embeddedSoundDeck}. */
const WAV = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA=';

/**
 * A real two-slide archive whose second slide has an embedded transition sound
 * and an on-click effect with an embedded sound, for the tests that check every
 * binding can play a deck's embedded sounds.
 */
export async function embeddedSoundDeck(Handler) {
	const { handler, createSlide } = await Handler.create({ initialSlideCount: 0 });
	try {
		const first = createSlide('Blank')
			.addText('Embedded sounds', { x: 50, y: 50, width: 400, height: 60 })
			.build();
		const second = createSlide('Blank')
			.addText('Animated', { x: 50, y: 50, width: 400, height: 60 })
			.build();
		second.transition = { type: 'fade', durationMs: 300, soundData: WAV };
		second.animations = [
			{
				id: 'effect',
				elementId: second.elements[0].id,
				entrance: 'appear',
				trigger: 'onClick',
				soundData: WAV,
			},
		];
		return await handler.save([first, second]);
	} finally {
		handler.dispose();
	}
}
