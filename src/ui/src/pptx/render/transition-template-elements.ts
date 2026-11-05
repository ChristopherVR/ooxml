/**
 * Which layout and master shapes each layer of a slide transition paints.
 *
 * A transition paints the leaving slide over the arriving one, and each of
 * them belongs behind its OWN layout. The viewer's "current" template is the
 * active slide's, which during a transition is already the arriving slide, so
 * handing it to the leaving layer repainted that slide over the next slide's
 * layout: a layout background picture vanished the moment a cover began
 * (issue #141, slides 1 to 2).
 *
 * @module render/transition-template-elements
 */
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

import { visibleTemplateElements } from './template-editing';

/** The template shapes behind each layer of a transition. */
export interface TransitionTemplateElements {
	/** Behind the leaving slide: its own layout and master. */
	readonly outgoing: PptxElement[];
	/** Behind the arriving slide: its own layout and master. */
	readonly incoming: PptxElement[];
}

type TemplateSlide = Pick<PptxSlide, 'id' | 'showMasterShapes'>;

/**
 * Each transition layer's template shapes, looked up by that layer's own
 * slide and filtered the way every other template render site filters them
 * ("Hide Background Graphics" hides them; see `visibleTemplateElements`).
 *
 * @param outgoingSlide - The slide the transition leaves.
 * @param incomingSlide - The slide it arrives on, when there is one.
 * @param templateElementsBySlideId - The loaded template shapes, by slide id.
 */
export function transitionTemplateElements(
	outgoingSlide: TemplateSlide,
	incomingSlide: TemplateSlide | undefined,
	templateElementsBySlideId: Readonly<Record<string, readonly PptxElement[]>>,
): TransitionTemplateElements {
	const of = (slide: TemplateSlide | undefined): PptxElement[] =>
		slide ? [...visibleTemplateElements(slide, templateElementsBySlideId[slide.id] ?? [])] : [];
	return { outgoing: of(outgoingSlide), incoming: of(incomingSlide) };
}
