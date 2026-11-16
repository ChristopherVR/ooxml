import { ELEMENT_ALT_TEXT_SECTIONS } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

export interface AccessibilitySection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * The Accessibility section: the Alt Text fields (description and title) of a
 * shape, text box, connector or graphic frame, the shared
 * `ELEMENT_ALT_TEXT_SECTIONS` drawn by the generic renderer. A picture's alt
 * text is drawn in `image-section.ts` from the same descriptors, so this
 * section is gated on `state.showAccessibilitySection` (shared's
 * `shouldShowAccessibilitySection`).
 */
export function createAccessibilitySection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
): AccessibilitySection {
	const el = section(t('pptx.accessibility.heading'));
	// Scroll target for the "Edit Alt Text" element context-menu command; see
	// `ooxml-ui/pptx`'s `context-menu-inspector-anchor.ts`.
	el.setAttribute('data-pptx-inspector-section', 'alt-text');
	const fields = createInspectorFields(doc, t, {
		sections: ELEMENT_ALT_TEXT_SECTIONS,
		onPatch: (patch) => handlers.patchSelectedElement(patch),
	});
	el.appendChild(fields.el);

	return {
		el,
		update(state) {
			const element = state.selectedElement;
			el.hidden = !state.hasSelection || !state.showAccessibilitySection || !element;
			if (element) {
				fields.update(element, !state.showAccessibilitySection, element.id);
			}
		},
	};
}
