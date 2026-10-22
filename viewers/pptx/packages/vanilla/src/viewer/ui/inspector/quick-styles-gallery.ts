import {
	buildQuickStylesGallery,
	quickStyleShapeStyle,
	registerPptxWebControls,
} from 'ooxml-ui/pptx';
import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import type { InspectorHandlers, InspectorState } from './types';

/**
 * The Quick Styles gallery (React's `QuickStylesGallery`): PowerPoint's Shape
 * Styles grid, one click applying a whole fill/stroke/shadow preset to the
 * selection. It is the shared gallery (`pptx-ui-ribbon-gallery`, a compact
 * panel) over the shared presets, so every binding offers the same tiles in
 * the same order; a pick hands the preset's style to `setShapeStyle`, which
 * merges it over the shape's own. The section's heading is its caption, so a
 * click on it never reaches a tile (as a wrapping `<label>` would).
 */
export function createQuickStylesGallery(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: Pick<InspectorHandlers, 'setShapeStyle'>,
) {
	registerPptxWebControls();
	const el = section(t('pptx.shape.quickStyles'));
	const gallery = doc.createElement('pptx-ui-ribbon-gallery');
	gallery.setAttribute('mode', 'panel');
	gallery.setAttribute('compact', '');
	gallery.translateLabel = (key, params) => t(key, params ? { ...params } : undefined);
	gallery.descriptor = buildQuickStylesGallery();
	gallery.addEventListener('gallery-pick', (event) => {
		const style = quickStyleShapeStyle((event as RibbonGalleryPickEvent).detail.itemId);
		if (style) {
			handlers.setShapeStyle(style);
		}
	});
	el.appendChild(gallery);

	return {
		el,
		update(state: InspectorState) {
			// Same gate as the Fill & Stroke section: a quick style is a shape
			// style, so it is meaningless for a picture or a chart.
			el.hidden = !state.canShape;
			gallery.disabled = !state.canShape;
		},
	};
}
