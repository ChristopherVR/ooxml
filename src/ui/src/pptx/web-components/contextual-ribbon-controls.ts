import { html, nothing } from 'lit';
import { renderMediaControls } from './contextual-media-ribbon';
import { renderSmartArtControls } from './contextual-smartart-ribbon';
import { ribbonFields } from './contextual-ribbon-fields';
import type { ContextualControls } from './contextual-ribbon-fields';

export type { ContextualControls };

export function renderContextualControls(context: ContextualControls) {
	const element = context.state.element;
	const fields = ribbonFields(context);
	const controls =
		element?.type === 'media'
			? renderMediaControls(context, fields, element)
			: element?.type === 'smartArt'
				? renderSmartArtControls(context, fields, element)
				: nothing;
	return html`<div class="contextual-fields">${controls}</div>`;
}
