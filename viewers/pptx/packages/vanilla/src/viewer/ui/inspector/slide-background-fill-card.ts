import {
	SLIDE_BACKGROUND_SECTIONS,
	slideBackgroundClearPatch,
	slideBackgroundClearState,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import type { DeckCard } from './deck-card-helpers';
import { makeDeckButton, makeSection } from './deck-card-helpers';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers } from './types';

/**
 * The BACKGROUND card: the shared `SLIDE_BACKGROUND_SECTIONS` (the active
 * slide's solid colour and Hide background graphics) drawn by the generic
 * renderer, plus the Clear Background action. Clear visibility and
 * enablement come from the shared `slideBackgroundClearState`; every change is
 * one `updateActiveSlide` patch, so it is one undo step.
 */
export function createSlideBackgroundFillCard(
	doc: Document,
	t: Translator,
	handlers: Pick<InspectorHandlers, 'updateActiveSlide' | 'pushRecentColor'>,
): DeckCard {
	const { el, body } = makeSection(doc, t('pptx.viewer.background'));
	const fields = createInspectorFields(doc, t, {
		sections: SLIDE_BACKGROUND_SECTIONS,
		onPatch: (patch) => handlers.updateActiveSlide(patch),
		onColorCommit: (hex) => handlers.pushRecentColor(hex),
	});
	const clear = makeDeckButton(doc, t('pptx.slideBackground.clearBackground'), () =>
		handlers.updateActiveSlide(slideBackgroundClearPatch()),
	);
	body.append(fields.el, clear);
	return {
		el,
		update(state) {
			el.hidden = !state.activeSlide;
			fields.update(state.activeSlide, !state.editable, state.activeSlide?.id);
			const clearState = slideBackgroundClearState(state.activeSlide, state.editable);
			clear.hidden = !clearState.visible;
			clear.disabled = !clearState.enabled;
		},
	};
}
