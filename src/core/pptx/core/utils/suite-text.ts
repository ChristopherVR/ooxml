import { PptxHandler } from '../PptxHandler';

/** Bounded text-only inspection for hosts; binary media never enters the AI context. */
export async function inspectPresentationText(bytes: Uint8Array) {
	const data = await new PptxHandler().load(new Uint8Array(bytes).buffer);
	return {
		slides: data.slides.map((slide, index) => ({
			index,
			elements: slide.elements.flatMap((element) =>
				'text' in element || 'textSegments' in element
					? [
							{
								id: element.id,
								type: element.type,
								text:
									element.text ??
									element.textSegments?.map((segment) => segment.text).join('') ??
									'',
								runs: element.textSegments?.map((segment, runIndex) => ({
									runIndex,
									text: segment.text,
								})),
							},
						]
					: [],
			),
		})),
	};
}

/** Change one existing run and preserve all other runs, styles and document parts. */
export async function setPresentationRunText(
	bytes: Uint8Array,
	slideIndex: number,
	elementId: string,
	runIndex: number,
	text: string,
) {
	if (
		!Number.isSafeInteger(slideIndex) ||
		slideIndex < 0 ||
		!Number.isSafeInteger(runIndex) ||
		runIndex < 0
	)
		throw new Error('Invalid slide or run index');
	const handler = new PptxHandler();
	const data = await handler.load(new Uint8Array(bytes).buffer);
	const element = data.slides[slideIndex]?.elements.find((item) => item.id === elementId);
	if (!element || !('textSegments' in element) || !element.textSegments?.[runIndex])
		throw new Error('Only existing top-level text runs can be edited');
	if (element.textSegments[runIndex].fieldType || element.textSegments[runIndex].equationXml)
		throw new Error('Fields and equations must be edited in the presentation editor');
	element.textSegments[runIndex].text = text;
	element.text = element.textSegments.map((segment) => segment.text).join('');
	return handler.save(data.slides);
}

export interface PresentationOutlineSlide {
	title: string;
	/** Body paragraphs, drawn as a bulleted list under the title. */
	bullets?: readonly string[];
	notes?: string;
}

/**
 * A new 16:9 deck from slide titles and bullet points: a title slide when the first entry has no
 * bullets, then one title-and-content slide per entry. Text boxes stay ordinary editable text.
 */
export async function createPresentationFromOutline(
	slides: readonly PresentationOutlineSlide[],
	options: { title?: string } = {},
): Promise<Uint8Array> {
	// Hosts pass parsed JSON, so the shape is checked at run time too.
	if (!(slides instanceof Array) || !slides.length || slides.length > 200)
		throw new Error('A presentation outline has 1 to 200 slides');
	const { PresentationBuilder } = await import('../builders/sdk/PresentationBuilder');
	const { handler, data, createSlide } = await PresentationBuilder.create({
		...(options.title ? { title: options.title } : {}),
	});
	const plain = (value: unknown, max = 2000) => {
		if (typeof value !== 'string' || value.length > max) throw new Error('Invalid slide text');
		return value;
	};
	slides.forEach((slide, index) => {
		const bullets = (slide.bullets ?? []).map((item) => plain(item));
		const cover = index === 0 && !bullets.length;
		const builder = createSlide(cover ? 'Title Slide' : 'Title and Content').addText(
			plain(slide.title, 500),
			cover
				? { x: 96, y: 250, width: 1088, height: 160, fontSize: 60, alignment: 'center' }
				: { x: 64, y: 40, width: 1152, height: 110, fontSize: 44 },
		);
		if (bullets.length)
			builder.addText(bullets.map((text) => text.replace(/\n/g, ' ')).join('\n'), {
				x: 64,
				y: 160,
				width: 1152,
				height: 500,
				fontSize: 28,
			});
		if (slide.notes) builder.setNotes(plain(slide.notes, 20_000));
		const built = builder.build();
		// One bullet per body paragraph: the builder splits the body text on its line breaks.
		const body = bullets.length ? built.elements[built.elements.length - 1] : undefined;
		if (body?.type === 'text')
			for (const [i, segment] of (body.textSegments ?? []).entries())
				if (!segment.isParagraphBreak && (i === 0 || body.textSegments?.[i - 1]?.isParagraphBreak))
					segment.bulletInfo = { char: '•' };
		data.slides.push(built);
	});
	return handler.save(data.slides);
}
