/**
 * The public contract of `SlidesPane`, the slide thumbnail rail a host places
 * beside its own `SlideCanvas` in a custom shell, and `createSlidesPaneProps`,
 * which wires it to `createViewerState`'s bag the same way `ViewerBody` wires
 * the `ThumbnailRail` that `PowerPointViewer` renders.
 *
 * The one difference is the slide context menu's "Layout" entry: in
 * `PowerPointViewer` it opens the Layout gallery, which a custom shell renders
 * itself (or not at all). Here it makes the slide active and hands the request
 * to the host's `onopenlayoutgallery`.
 */
import { buildSlidesPaneFieldContext, buildSlidesPaneTableStyleContext } from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	FieldSubstitutionContext,
	SlidesPaneLayoutGalleryRequest,
} from 'ooxml-ui/pptx';
import type { ParsedTableStyleMap, PptxSection, PptxSlide, PptxTheme } from 'pptx-viewer-core';

import type { ViewerStateBag } from '../state/create-viewer-state-types';

export type { SlidesPaneLayoutGalleryRequest };

/**
 * The theme and table style map the thumbnails resolve table colours from,
 * as returned by `buildSlidesPaneTableStyleContext`.
 */
export interface SlidesPaneTableStyleContext {
	theme: PptxTheme | undefined;
	tableStyleMap: ParsedTableStyleMap | undefined;
}

/** Props of the standalone `SlidesPane` (the slide thumbnail rail). */
export interface SlidesPaneProps {
	/**
	 * The slides as the canvas renders them, with their layout and master
	 * elements merged in (`ViewerStateBag.displaySlides`), so each thumbnail
	 * draws what the canvas draws.
	 */
	slides: PptxSlide[];
	/** Zero-based index of the slide shown on the canvas; its row is marked `aria-current`. */
	activeSlideIndex: number;
	/** The deck's slide size in CSS px; thumbnails are scaled down from it. */
	canvasSize: CanvasSize;
	/** Resolved media data URLs (`PresentationLoader.mediaDataUrls`) for images and video posters. */
	mediaDataUrls: Map<string, string>;
	/** The deck's sections; with any, the rail groups thumbnails under collapsible headers. */
	sections?: PptxSection[];
	/**
	 * Whether the rail edits the deck: drag reordering, the context menus, the
	 * Enter shortcut and the "New slide" footer. Default `false` (read-only, no footer).
	 */
	canEdit?: boolean;
	/**
	 * Deck-wide field context (`buildSlidesPaneFieldContext`): date/time,
	 * header/footer, document properties and the deck start number
	 * (`p:presentation/@firstSlideNum`). Each thumbnail re-points the slide
	 * number and title at its own slide. When omitted, the rail uses the context
	 * of the enclosing `createViewerState`, if any.
	 */
	fieldContext?: FieldSubstitutionContext;
	/**
	 * Theme and table style map for table thumbnails
	 * (`buildSlidesPaneTableStyleContext`). When omitted, the rail uses the
	 * context of the enclosing `createViewerState`, if any.
	 */
	tableStyleContext?: SlidesPaneTableStyleContext;
	/** A thumbnail was clicked: show that slide. Selecting never edits the deck. */
	onselectslide: (index: number) => void;
	/** A thumbnail was dragged from `fromIndex` and dropped on `toIndex`. */
	onmoveslide?: (fromIndex: number, toIndex: number) => void;
	/** The footer's "New slide" button. The footer renders only with `canEdit` and this callback. */
	onaddslide?: () => void;
	/** Insert a slide after `index` (the context menu's New Slide, and Enter on a thumbnail). */
	onaddslideafter?: (index: number) => void;
	/** Duplicate every selected slide (multi-selection with Ctrl/Shift click). */
	onduplicateslides?: (indexes: number[]) => void;
	/** Delete every selected slide. */
	ondeleteslides?: (indexes: number[]) => void;
	/** Toggle hidden on every selected slide. */
	onhideslides?: (indexes: number[]) => void;
	/**
	 * The context menu's "Layout" entry. The rail first makes the slide active
	 * through `onselectslide`, then calls this with the slide and the click
	 * point, so a host can open its own Layout gallery there.
	 */
	onopenlayoutgallery?: (request: SlidesPaneLayoutGalleryRequest) => void;
	/** Add a section starting after slide `index` (context menus' Add Section). */
	onaddsection?: (index: number) => void;
	/** Rename a section (double-click or the section menu's Rename). */
	onrenamesection?: (sectionId: string, name: string) => void;
	/** Delete a section; its slides stay in the deck. */
	ondeletesection?: (sectionId: string) => void;
	/** Move a section, with its slides, one place up. */
	onmovesectionup?: (sectionId: string) => void;
	/** Move a section, with its slides, one place down. */
	onmovesectiondown?: (sectionId: string) => void;
	/** Collapse or expand a section's thumbnails. */
	ontogglesection?: (sectionId: string) => void;
}

/** Options of {@link createSlidesPaneProps}. */
export interface CreateSlidesPanePropsOptions {
	/** See {@link SlidesPaneProps.onopenlayoutgallery}. */
	onopenlayoutgallery?: (request: SlidesPaneLayoutGalleryRequest) => void;
}

/**
 * `SlidesPane` props over a `createViewerState` bag, wired like the full
 * viewer's rail: the same merged slides, sections, slide and section
 * operations (so edits share the undo history and collaboration sync) and the
 * canvas's field and table contexts. Every value is a live getter, so spread
 * the result as is: `<SlidesPane {...createSlidesPaneProps(state)} />`.
 */
export function createSlidesPaneProps(
	state: ViewerStateBag,
	options: CreateSlidesPanePropsOptions = {},
): SlidesPaneProps {
	const { editor } = state;
	const select = (index: number): void => state.viewer.goTo(index);
	const selectInserted = (index: number | null): void => {
		if (index !== null) {
			select(index);
		}
	};
	return {
		get slides() {
			return state.displaySlides;
		},
		get activeSlideIndex() {
			return state.viewer.current;
		},
		get canvasSize() {
			return state.loader.canvasSize;
		},
		get mediaDataUrls() {
			return state.loader.mediaDataUrls;
		},
		get sections() {
			return editor.sections;
		},
		get canEdit() {
			return state.editingActive;
		},
		get fieldContext() {
			return buildSlidesPaneFieldContext({
				headerFooter: editor.headerFooter,
				customProperties: editor.customProperties,
				firstSlideNumber: state.loader.firstSlideNumber,
			});
		},
		get tableStyleContext() {
			return buildSlidesPaneTableStyleContext(
				state.loader.presentationTheme,
				state.loader.tableStyleMap,
			);
		},
		get onopenlayoutgallery() {
			return options.onopenlayoutgallery;
		},
		onselectslide: select,
		onmoveslide: (fromIndex, toIndex) =>
			selectInserted(editor.slidesOps.moveSlide(fromIndex, toIndex)),
		onaddslide: () => selectInserted(editor.slidesOps.insertSlideAfterCurrent()),
		onaddslideafter: (index) => selectInserted(editor.slidesOps.insertSlideAfter(index)),
		onduplicateslides: (indexes) => editor.slidesOps.duplicateSlides(indexes),
		ondeleteslides: (indexes) => editor.slidesOps.deleteSlides(indexes),
		onhideslides: (indexes) => editor.slidesOps.toggleSlidesHidden(indexes),
		onaddsection: (index) => {
			editor.sectionOps.add(state.t('pptx.sections.defaultName'), index);
		},
		onrenamesection: (sectionId, name) => editor.sectionOps.rename(sectionId, name),
		ondeletesection: (sectionId) => editor.sectionOps.delete(sectionId),
		onmovesectionup: (sectionId) => editor.sectionOps.moveUp(sectionId),
		onmovesectiondown: (sectionId) => editor.sectionOps.moveDown(sectionId),
		ontogglesection: (sectionId) => editor.sectionOps.toggle(sectionId),
	};
}
