import {
	registerPptxWebControls,
	schemaLabel,
	SMARTART_COLOR_SCHEME_LABEL_KEYS,
	SMARTART_LAYOUT_LABEL_KEYS,
} from 'ooxml-ui/pptx';
import type { PptxUiSmartArtTextPaneElement, SmartArtTextPaneChangeEvent } from 'ooxml-ui/pptx';
import type { SmartArtColorScheme, SmartArtLayoutType } from 'pptx-viewer-core';
import { SWITCHABLE_LAYOUT_TYPES } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeSelectField } from './controls-extra';
import type { InspectorHandlers, InspectorState } from './types';

const COLOR_SCHEMES: readonly SmartArtColorScheme[] = [
	'colorful1',
	'colorful2',
	'colorful3',
	'monochromatic1',
	'monochromatic2',
];

export interface SmartArtSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * SmartArt layout and colour controls aligned with React, and the text pane, which is the shared
 * `pptx-ui-smartart-text-pane` element (outline with levels, keyboard structure edits, move,
 * sub-item, remove, per-node style). This section feeds it the data and commits what it emits.
 */
export function createSmartArtSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
): SmartArtSection {
	const el = section(t('pptx.smartart.title'));
	const layoutLabel = createEl(doc, 'span', 'pptxv-smartart-label');
	layoutLabel.textContent = t('pptx.smartart.switchLayout');
	el.appendChild(layoutLabel);

	const layoutGrid = createEl(doc, 'div', 'pptxv-smartart-layout-grid');
	const layoutButtons = new Map<SmartArtLayoutType, HTMLButtonElement>();
	for (const layout of SWITCHABLE_LAYOUT_TYPES) {
		const button = createEl(doc, 'button', 'pptxv-smartart-layout-button');
		button.type = 'button';
		button.dataset.testid = `smartart-layout-${layout}`;
		// Same key the hand-built template produced, but routed through the shared
		// map so a layout family missing from the dictionary shows its own name
		// rather than a de-camel-cased key tail.
		button.textContent = schemaLabel(SMARTART_LAYOUT_LABEL_KEYS, layout, t);
		button.addEventListener('click', () => handlers.setSmartArtLayout(layout));
		layoutButtons.set(layout, button);
		layoutGrid.appendChild(button);
	}
	el.appendChild(layoutGrid);

	const colorScheme = makeSelectField(doc, {
		label: t('pptx.smartart.colorScheme'),
		// The `dgm:colorsDef` family name is what core stores, so it stays the
		// option value; the caption is spelled from the shared map instead of
		// showing the user `monochromatic2`.
		options: COLOR_SCHEMES.map((scheme) => ({
			value: scheme,
			label: schemaLabel(SMARTART_COLOR_SCHEME_LABEL_KEYS, scheme, t),
		})),
		onChange: handlers.setSmartArtColorScheme,
	});
	const select = colorScheme.el.querySelector<HTMLElement>('pptx-ui-select');
	if (select) {
		select.dataset.testid = 'smartart-color-scheme';
	}
	el.appendChild(colorScheme.el);

	registerPptxWebControls();
	const pane = doc.createElement('pptx-ui-smartart-text-pane') as PptxUiSmartArtTextPaneElement;
	pane.dataset.testid = 'smartart-text-pane';
	pane.addEventListener('pptx-smartart-data-change', (event) => {
		const { smartArtData, recentColor } = (event as SmartArtTextPaneChangeEvent).detail;
		handlers.replaceSmartArtData(smartArtData);
		if (recentColor) {
			handlers.pushRecentColor(recentColor);
		}
	});
	el.appendChild(pane);

	return {
		el,
		update(state) {
			el.hidden = !state.isSmartArt;
			const current = state.smartArtData?.resolvedLayoutType;
			for (const [layout, button] of layoutButtons) {
				button.disabled = !state.isSmartArt;
				button.classList.toggle('is-active', layout === current);
				button.setAttribute('aria-pressed', String(layout === current));
			}
			colorScheme.setValue(state.smartArtData?.colorScheme ?? 'colorful1');
			colorScheme.setDisabled(!state.isSmartArt);
			pane.state = {
				smartArtData: state.isSmartArt ? state.smartArtData : undefined,
				canEdit: true,
				translate: t,
			};
		},
	};
}
