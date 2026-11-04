import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import type { SlidesPaneDeckSources } from '../composables/slides-pane-contexts';
import type { CanvasSize } from '../types';

/**
 * Props of `SlideThumbnail`, a static, scaled-down drawing of one slide as the slides pane
 * shows it. It renders from the slide model, so it follows unsaved edits.
 *
 * Feed it from the same values a host passes to `SlidesPane`: `templateElementsBySlideId[slide.id]`,
 * `canvasSize`, `mediaDataUrls`, `headerFooter`, `customProperties`, `firstSlideNumber`, `locale`,
 * `theme` and `tableStyleMap`.
 */
export interface SlideThumbnailProps extends SlidesPaneDeckSources {
	slide: PptxSlide;
	/** The slide's inherited layout and master elements, drawn beneath its own. Defaults to none. */
	templateElements?: PptxElement[];
	/** Slide size in px, for the thumbnail aspect ratio and scale. */
	canvasSize: CanvasSize;
	/**
	 * Thumbnail width in pixels; the height follows the deck's aspect ratio. Defaults to the
	 * slides pane thumbnail width.
	 */
	previewWidth?: number;
	/** Media data URLs by part path (`useLoadContent().mediaDataUrls`), for pictures and posters. */
	mediaDataUrls?: Map<string, string>;
}
