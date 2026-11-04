import type { CanvasSize, FieldSubstitutionContext } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import type { SlidesPaneTableStyleContext } from './slides-pane-props';

/**
 * Props of `SlideThumbnail`, a static, scaled-down drawing of one slide as `SlidesPane` shows
 * it. It renders from the slide model, so it follows unsaved edits.
 *
 * Feed it from the same values as `SlidesPane` (for example `createSlidesPaneProps(state)`):
 * a slide from `slides`, plus `canvasSize`, `mediaDataUrls`, `fieldContext` and
 * `tableStyleContext`. Those `slides` already carry their layout and master elements, so leave
 * `templateElements` empty for them; pass it only for a slide whose template layer is separate.
 */
export interface SlideThumbnailProps {
	slide: PptxSlide;
	/** The slide's inherited layout and master elements, drawn beneath its own. Default none. */
	templateElements?: PptxElement[];
	/** The deck's slide size in CSS px; the thumbnail is scaled down from it. */
	canvasSize: CanvasSize;
	/**
	 * Thumbnail width in pixels; the height follows the deck's aspect ratio. Defaults to the
	 * slides pane thumbnail width.
	 */
	previewWidth?: number;
	/** Resolved media data URLs (`PresentationLoader.mediaDataUrls`) for images and video posters. */
	mediaDataUrls?: Map<string, string>;
	/** Deck-wide field context; see `SlidesPaneProps.fieldContext`. */
	fieldContext?: FieldSubstitutionContext;
	/** Theme and table style map for tables; see `SlidesPaneProps.tableStyleContext`. */
	tableStyleContext?: SlidesPaneTableStyleContext;
}
