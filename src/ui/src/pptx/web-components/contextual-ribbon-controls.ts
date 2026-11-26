import { html, nothing } from 'lit';
import { canInteractWithElement } from '../render/element-locks';
import { quarterTurnPatch } from '../render/ribbon-drawing-arrange';
import { renderMediaControls } from './contextual-media-ribbon';
import { renderSmartArtControls } from './contextual-smartart-ribbon';
import { renderSmartArtDesignControls } from './contextual-smartart-design';
import { ribbonFields } from './contextual-ribbon-fields';
import type { ContextualControls } from './contextual-ribbon-fields';

export type { ContextualControls };

export function renderContextualControls(context: ContextualControls) {
	const element = context.state.element;
	const fields = ribbonFields(context);
	if (context.group === 'shapeFormat.arrange' && element) {
		const { text, command, patch, unavailable } = fields;
		const reason = canInteractWithElement(element, 'rotate') ? undefined : unavailable;
		const selectionPane = context.state.selectionPane;
		return html`<div class="contextual-fields">
			${command('rotateRight', text('pptx.pictureFormat.rotateRight', 'Rotate Right 90°'), () => patch(quarterTurnPatch(element, true)), reason)}
			${command('rotateLeft', text('pptx.pictureFormat.rotateLeft', 'Rotate Left 90°'), () => patch(quarterTurnPatch(element, false)), reason)}
			${selectionPane ? command('selectionPane', text('pptx.selectionPane.title', 'Selection Pane'), selectionPane) : nothing}
		</div>`;
	}
	const controls =
		element?.type === 'media'
			? renderMediaControls(context, fields, element)
			: element?.type === 'smartArt'
				? context.group === 'smartArtDesign.createGraphic'
					? renderSmartArtDesignControls(context, fields, element)
					: renderSmartArtControls(context, fields, element)
				: nothing;
	return html`<div class="contextual-fields">${controls}</div>`;
}
