import { ELEMENT_GEOMETRY_SECTIONS } from 'ooxml-ui/pptx';

import type { GeometryPatch } from '../../editor/editor-edit-ops';
import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeButton } from '../controls';
import { createIcon } from '../icons';
import { createInspectorFields } from './inspector-fields';
import type { InspectorState } from './types';

export interface PositionSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * The universal Position & Size section: the shared X/Y/W/H/rotation
 * descriptors (`ELEMENT_GEOMETRY_SECTIONS`) drawn by the generic renderer, in
 * the panel's two-column grid, plus the element lock toggle.
 */
export function createPositionSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	setGeometry: (patch: GeometryPatch) => void,
	toggleElementLock: () => void,
): PositionSection {
	const el = section(t('pptx.arrange.positionSize'));
	// Scroll target for the "Size and Position" element context-menu command;
	// see `pptx-viewer-shared`'s `context-menu-inspector-anchor.ts`.
	el.setAttribute('data-pptx-inspector-section', 'transform');

	// Move the section's own caption into a header row alongside the lock
	// toggle, matching React's ElementInspectorBody / Vue's ArrangePanel.
	const caption = el.querySelector<HTMLElement>('.pptxv-inspector-section-title');
	const headerRow = createEl(doc, 'div', 'pptxv-inspector-lock-row');
	if (caption) {
		el.insertBefore(headerRow, caption);
		headerRow.appendChild(caption);
	} else {
		el.insertBefore(headerRow, el.firstChild);
	}
	// Shared decides both what reads as "locked" and what the toggle writes, so
	// the button's state can never drift from what the canvas enforces.
	const lockBtn = makeButton(doc, {
		label: t('pptx.inspector.lock'),
		icon: 'lock-open',
		className: 'pptxv-inspector-lock-btn',
		onClick: toggleElementLock,
	});
	headerRow.appendChild(lockBtn.btn);

	const geometry = createInspectorFields(doc, t, {
		sections: ELEMENT_GEOMETRY_SECTIONS,
		onPatch: (patch) => setGeometry(patch as GeometryPatch),
		sectionClassName: 'pptxv-inspector-grid',
	});
	el.appendChild(geometry.el);

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection;
			geometry.update(
				{
					x: state.x,
					y: state.y,
					width: state.width,
					height: state.height,
					rotation: state.rotation,
				},
				!state.hasSelection,
			);
			lockBtn.setActive(state.isLocked);
			lockBtn.setDisabled(!state.hasSelection);
			const label = t(state.isLocked ? 'pptx.inspector.unlock' : 'pptx.inspector.lock');
			lockBtn.btn.title = label;
			lockBtn.btn.setAttribute('aria-label', label);
			lockBtn.btn.replaceChildren(createIcon(doc, state.isLocked ? 'lock' : 'lock-open'));
		},
	};
}
