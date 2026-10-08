import { defineRibbonActions, type OfficeUiRibbonActions } from '../../ribbon/ribbon-actions';
import type { TabRowActionsViewState } from '../render';
import { attachControlStyles } from './control-styles';
import { bridgeCss } from './office-token-bridge';

/**
 * `pptx-ui-ribbon-actions`: the shared `office-ui-ribbon-actions` (Comments and Share at the right
 * end of the ribbon tab row, as Word and Excel draw them) under the pptx tag, with the pptx token
 * bridge. Controlled through `state` (`buildTabRowActionsState`); events `comments-toggle` and
 * `share-request`.
 */
export function definePptxRibbonActions(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-ribbon-actions')) {
		return;
	}
	defineRibbonActions(registry);
	const Base = registry.get('office-ui-ribbon-actions') as unknown as typeof OfficeUiRibbonActions;
	class PptxRibbonActions extends Base {
		static override commentsEvent = 'comments-toggle';
		static override shareEvent = 'share-request';
		#model: TabRowActionsViewState | undefined;
		constructor() {
			super();
			if (this.shadowRoot) {
				attachControlStyles(this.shadowRoot, bridgeCss('pptx-ui-ribbon-actions'));
			}
		}
		get state(): TabRowActionsViewState | undefined {
			return this.#model;
		}
		set state(value: TabRowActionsViewState | undefined) {
			this.#model = value;
			if (value) {
				Object.assign(this, value);
			}
		}
	}
	registry.define(
		'pptx-ui-ribbon-actions',
		PptxRibbonActions as unknown as CustomElementConstructor,
	);
}
