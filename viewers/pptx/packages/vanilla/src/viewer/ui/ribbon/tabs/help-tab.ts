import { buildHelpRibbon, EMPTY_RESOLVED_CUSTOMIZATION, isDialogAvailable } from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import { createEl } from '../../../render';
import type { RibbonNavHandlers } from '../ribbon-types';

/** DOM adapter for the shared Help tab. */
export function createHelpTab(
	doc: Document,
	t: Translator,
	handlers: RibbonNavHandlers,
): HTMLElement {
	const el = createEl(doc, 'div', 'pptxv-ribbon-tab-content');
	const section = doc.createElement('pptx-ui-ribbon-section');
	const options = isDialogAvailable(
		handlers.getCustomization?.() ?? EMPTY_RESOLVED_CUSTOMIZATION,
		'options',
	);
	// Keyboard Shortcuts is a tab of Options in this binding.
	section.groups = buildHelpRibbon(t, { options, shortcuts: options });
	section.addEventListener('command-request', (event) => {
		switch ((event as RibbonCommandRequestEvent).detail.id) {
			case 'help.help.options':
				handlers.openSettings('general');
				break;
			case 'help.help.keyboardShortcuts':
				handlers.openSettings('shortcuts');
				break;
			case 'help.help.accessibility':
				handlers.openAccessibility();
				break;
		}
	});
	el.append(section);
	return el;
}
