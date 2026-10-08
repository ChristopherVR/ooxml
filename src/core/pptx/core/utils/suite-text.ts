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
