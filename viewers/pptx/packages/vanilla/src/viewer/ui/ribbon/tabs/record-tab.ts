import { buildRecordRibbon, runRecordRibbonAction } from 'ooxml-ui/pptx';
import type { RecordTarget, RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import { createEl } from '../../../render';
import type { RibbonSlideShowHandlers } from '../ribbon-types';

/** Keep Vanilla's existing rehearsal workflow behind the shared Record view. */
export function createRecordTab(
	doc: Document,
	t: Translator,
	handlers: RibbonSlideShowHandlers,
	getTarget?: () => RecordTarget | undefined,
): HTMLElement & { sync(editable: boolean): void } {
	const el = createEl(doc, 'div', 'pptxv-ribbon-tab-content');
	const section = doc.createElement('pptx-ui-ribbon-section');
	section.groups = buildRecordRibbon(t);
	let canEdit = false;
	const target = () => (canEdit ? getTarget?.() : undefined);
	section.addEventListener('command-request', (event) => {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (id === 'record.record.fromBeginning' || id === 'record.record.fromCurrent') {
			handlers.startRehearsal();
		} else runRecordRibbonAction(section, id, target, t);
	});
	el.append(section);
	return Object.assign(el, {
		sync(editable: boolean) {
			canEdit = editable;
			section.groups = buildRecordRibbon(t, target());
		},
	});
}
