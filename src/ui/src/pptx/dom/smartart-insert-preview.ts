import type { PptxSlide, SmartArtLayout } from 'ooxml-core/pptx';

import {
	buildSmartArtPreviewElement,
	SMARTART_PREVIEW_ELEMENT_HEIGHT,
	SMARTART_PREVIEW_ELEMENT_WIDTH,
} from '../render/preview-elements';
import type { Translator } from '../i18n/translator';
import { createEl } from './dom';
import { renderSmartArtSvg } from './elements/smartart';
import { createElementRendererRegistry } from './registry';
import type { ElementRenderContext } from './types';

/**
 * The live preview of an Insert SmartArt layout: the slide renderer's output for the exact
 * element the preset inserts (same layout, default items, colour scheme and style), scaled to
 * `size`, so the preview matches what lands on the slide. The PowerPoint Insert SmartArt dialog
 * (`pptx-ui-insert-smartart-dialog`) loads this module on first open and hands it to the shared
 * `office-ui-smartart-picker`. Moved from the vanilla binding
 * (`ui/ribbon/tabs/insert/smartart-gallery-preview.ts`), which every binding now shares.
 */
export function buildSmartArtInsertPreview(
	doc: Document,
	t: Translator,
	layout: SmartArtLayout,
	size: { width: number; height: number },
): HTMLElement {
	const scale = Math.min(
		size.width / SMARTART_PREVIEW_ELEMENT_WIDTH,
		size.height / SMARTART_PREVIEW_ELEMENT_HEIGHT,
	);
	const box = createEl(doc, 'div', 'pptxv-smartart-option-preview', {
		width: `${Math.round(SMARTART_PREVIEW_ELEMENT_WIDTH * scale)}px`,
		height: `${Math.round(SMARTART_PREVIEW_ELEMENT_HEIGHT * scale)}px`,
		overflow: 'hidden',
		pointerEvents: 'none',
		position: 'relative',
	});
	box.setAttribute('aria-hidden', 'true');
	const stage = createEl(doc, 'div', undefined, {
		width: `${SMARTART_PREVIEW_ELEMENT_WIDTH}px`,
		height: `${SMARTART_PREVIEW_ELEMENT_HEIGHT}px`,
		transform: `scale(${scale})`,
		transformOrigin: 'top left',
		position: 'relative',
	});
	box.appendChild(stage);

	const element = buildSmartArtPreviewElement(layout);
	const slide: PptxSlide = {
		id: 'smartart-preview',
		rId: 'smartart-preview',
		slideNumber: 1,
		elements: [element],
	};
	const context: ElementRenderContext = {
		document: doc,
		slide,
		canvasSize: { width: SMARTART_PREVIEW_ELEMENT_WIDTH, height: SMARTART_PREVIEW_ELEMENT_HEIGHT },
		scale: 1,
		mediaDataUrls: new Map(),
		t,
		smartArt3D: false,
		surfaceChart3D: false,
		barChart3D: false,
		lineChart3D: false,
		areaChart3D: false,
		pieChart3D: false,
		presenting: false,
		registry: createElementRendererRegistry(),
		renderElement: () => null,
	};
	const rendered = renderSmartArtSvg(element, 0, context);
	if (rendered) stage.appendChild(rendered);
	return box;
}
