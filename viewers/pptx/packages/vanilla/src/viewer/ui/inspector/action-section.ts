import { ACTION_SETTINGS_SECTIONS, actionSettingsTargetOf } from 'ooxml-ui/pptx';
import type { ActionSettingsChange } from 'ooxml-ui/pptx';
import type { ElementActionType, PptxElement } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

/** A named custom show, for the `customShow` action's picker. */
export interface ActionCustomShowOption {
	id: string;
	name: string;
}

/**
 * The Action Settings section: what the element does when clicked or hovered
 * in a slide show. Each trigger is the shared `ACTION_SETTINGS_SECTIONS` drawn
 * by the generic renderer over its context (`actionSettingsTargetOf`: its
 * action, the pending pick and the deck's slides and custom shows); a change
 * holds its type as pending and, once it can carry its target, commits
 * through `setElementAction`.
 */
export function createActionSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
) {
	const el = section(t('pptx.action.title'));
	// Neither the deck's slide count nor its custom shows are part of the
	// per-element inspector state, so the pane pushes both separately.
	let slideCount = 1;
	let customShows: readonly ActionCustomShowOption[] = [];
	let element: PptxElement | undefined;
	let disabled = true;
	const pending: Partial<Record<'click' | 'hover', ElementActionType>> = {};

	const editors = (['click', 'hover'] as const).map((trigger) => {
		const box = doc.createElement('fieldset');
		box.className = 'pptxv-action-fieldset';
		box.dataset['pptxActionTrigger'] = trigger;
		const legend = doc.createElement('legend');
		legend.textContent = t(trigger === 'click' ? 'pptx.action.onClick' : 'pptx.action.onHover');
		const fields = createInspectorFields(doc, t, {
			sections: ACTION_SETTINGS_SECTIONS,
			onPatch: (change: ActionSettingsChange) => {
				pending[trigger] = change.pendingType;
				if (change.commit) {
					handlers.setElementAction(trigger, change.commit);
				}
				paint();
			},
		});
		box.append(legend, fields.el);
		el.appendChild(box);
		return { trigger, fields };
	});

	function paint(): void {
		if (!element) {
			return;
		}
		for (const { trigger, fields } of editors) {
			fields.update(
				actionSettingsTargetOf(element, trigger, pending[trigger], slideCount, customShows),
				disabled,
				element.id,
			);
		}
	}

	return {
		el,
		setSlideCount(count: number) {
			slideCount = Math.max(1, count);
		},
		setCustomShows(shows: readonly ActionCustomShowOption[]) {
			customShows = shows;
		},
		update(state: InspectorState) {
			el.hidden = !state.hasSelection || !state.selectedElement;
			// A half-made pick belongs to the element it was made on.
			if (state.selectedElement?.id !== element?.id) {
				delete pending.click;
				delete pending.hover;
			}
			element = state.selectedElement;
			disabled = !state.hasSelection;
			paint();
		},
	};
}
