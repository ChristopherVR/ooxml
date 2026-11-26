import { html } from 'lit';
import type { SmartArtPptxElement } from 'ooxml-core/pptx';

import {
	applyRibbonGalleryItem,
	buildRibbonGallery,
} from '../render/ribbon-galleries/gallery-registry';
import { SMARTART_NODE_COMMANDS } from '../render/ribbon-galleries/smartart-commands';
import type { ContextualControls, ContextualFields } from './contextual-ribbon-fields';

export function renderSmartArtDesignControls(
	ctx: ContextualControls,
	fields: ContextualFields,
	element: SmartArtPptxElement,
) {
	const context = { element, selectedSmartArtNodeId: ctx.nodeIds.at(-1) ?? null };
	return SMARTART_NODE_COMMANDS.map(([control, gallery]) => {
		const descriptor = buildRibbonGallery(gallery, context);
		return html`<pptx-ui-ribbon-gallery
			data-ribbon-control=${`${ctx.group}.${control}`}
			mode="dropdown"
			.descriptor=${descriptor}
			.translateLabel=${ctx.state.translate}
			.disabled=${!fields.available}
			@gallery-pick=${() => {
				const result = applyRibbonGalleryItem(gallery, 'run', context);
				if (result?.kind === 'element') fields.patch(result.patch);
			}}
		></pptx-ui-ribbon-gallery>`;
	});
}
