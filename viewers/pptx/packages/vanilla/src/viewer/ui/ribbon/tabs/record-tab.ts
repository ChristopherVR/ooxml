import { buildRecordRibbon } from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import { createEl } from '../../../render';
import type { RibbonSlideShowHandlers } from '../ribbon-types';

/** Keep Vanilla's existing rehearsal workflow behind the shared Record view. */
export function createRecordTab(
	doc: Document,
	t: Translator,
	handlers: RibbonSlideShowHandlers,
): HTMLElement {
	const el = createEl(doc, 'div', 'pptxv-ribbon-tab-content');
	const section = doc.createElement('pptx-ui-ribbon-section');
	section.groups = buildRecordRibbon(t);
	// Both recording commands start the rehearsal; the placeholders are disabled.
	section.addEventListener('command-request', (event) => {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (id === 'record.record.fromBeginning' || id === 'record.record.fromCurrent') {
			handlers.startRehearsal();
		}
	});
	el.append(section);
	return el;
}
