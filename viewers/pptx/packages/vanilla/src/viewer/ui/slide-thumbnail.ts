/**
 * `renderSlideThumbnail`: one slide drawn the way the slides pane
 * (`createSlidesPane`) draws it, for a host that keeps its own thumbnail strip.
 *
 * @module ui/slide-thumbnail
 */
import {
	buildPreviewElements,
	buildSlidesPaneFieldContext,
	deriveSlideFieldContext,
	EDITOR_THUMBNAIL_WIDTH,
	slideThumbnailLayout,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import { createTranslator } from '../i18n';
import { createDefaultRegistry, createEl, renderSlideStage } from '../render';
import type { SlidesPaneOptions } from './slides-pane';

/**
 * Options of {@link renderSlideThumbnail}. Apart from `document`, `slide`,
 * `templateElements` and `previewWidth`, every field has the name, meaning and
 * default it has on {@link SlidesPaneOptions}, so a host can spread the options
 * it gives the slides pane.
 */
export interface SlideThumbnailOptions extends Pick<
	SlidesPaneOptions,
	| 'canvasSize'
	| 'mediaDataUrls'
	| 'headerFooter'
	| 'customProperties'
	| 'firstSlideNumber'
	| 'locale'
	| 'colorScheme'
	| 'fontScheme'
	| 'formatScheme'
	| 'tableStyleMap'
	| 't'
	| 'registry'
> {
	/** The document the thumbnail's nodes are created in. */
	document: Document;
	/** The slide to draw, as the host stores it (unsaved edits included). */
	slide: PptxSlide;
	/**
	 * The slide's inherited layout and master elements
	 * (`templateElementsBySlideId[slide.id]`), drawn beneath its own. Default none.
	 */
	templateElements?: readonly PptxElement[];
	/**
	 * Thumbnail width in CSS px; the height follows the deck's aspect ratio.
	 * Default: the slides pane's thumbnail width.
	 */
	previewWidth?: number;
	/** The deck's slides, for Zoom elements; the slides pane's `slides`. */
	slides?: SlidesPaneOptions['slides'];
}

/**
 * Draw `slide` as a static, scaled-down thumbnail and return its root box.
 *
 * The box fills its container's width and is `previewWidth` at the deck's
 * aspect ratio tall, with `overflow: hidden`. Inside, the slide's template
 * elements and its own elements are drawn, with the slide background, through
 * `renderSlideStage` at the full canvas size and scaled from the top-left,
 * exactly as the slides pane draws its thumbnails. The stage is static: it is
 * `inert` (no focus stops or pointer interaction), carries no editing contract,
 * and its elements carry no `data-element-id`, so element ids stay unique to
 * the editing canvas.
 *
 * It draws from the slide model, so call it again when the slide changes
 * (elements, background, hidden, `showMasterShapes`) and replace the old box:
 *
 * ```ts
 * const thumb = renderSlideThumbnail({
 * 	...paneOptions,
 * 	document,
 * 	slide,
 * 	templateElements: paneOptions.templateElementsBySlideId?.[slide.id],
 * 	previewWidth: 160,
 * });
 * row.replaceChildren(thumb);
 * ```
 *
 * Style it with the viewer stylesheet under a `.pptxv` ancestor, as the pane.
 */
export function renderSlideThumbnail(options: SlideThumbnailOptions): HTMLElement {
	const { document: doc, slide } = options;
	const canvasSize = {
		width: Math.max(1, options.canvasSize.width),
		height: Math.max(1, options.canvasSize.height),
	};
	const layout = slideThumbnailLayout(
		canvasSize.width,
		canvasSize.height,
		options.previewWidth ?? EDITOR_THUMBNAIL_WIDTH,
	);
	const deckFieldContext = buildSlidesPaneFieldContext({
		headerFooter: options.headerFooter,
		customProperties: options.customProperties,
		firstSlideNumber: options.firstSlideNumber,
		locale: options.locale,
	});
	const currentSlideIndex = options.slides?.indexOf(slide) ?? -1;
	const stage = renderSlideStage({
		document: doc,
		slide: { ...slide, elements: buildPreviewElements(slide, options.templateElements) },
		canvasSize,
		mediaDataUrls: options.mediaDataUrls ?? new Map(),
		colorScheme: options.colorScheme,
		fontScheme: options.fontScheme,
		formatScheme: options.formatScheme,
		tableStyleMap: options.tableStyleMap,
		fieldContext: deriveSlideFieldContext(deckFieldContext, slide),
		registry: options.registry ?? createDefaultRegistry(),
		t: options.t ?? createTranslator(),
		scale: layout.scale,
		slides: options.slides,
		currentSlideIndex: currentSlideIndex >= 0 ? currentSlideIndex : undefined,
	});
	stage.inert = true;
	for (const node of [stage, ...stage.querySelectorAll('[data-element-id]')]) {
		node.removeAttribute('data-element-id');
	}
	const box = createEl(doc, 'div', 'pptxv-slide-thumbnail', {
		position: 'relative',
		width: '100%',
		height: `${layout.height}px`,
		overflow: 'hidden',
	});
	box.appendChild(stage);
	return box;
}
