/**
 * `createSlidesPane`: the slide thumbnail rail as a standalone building block,
 * for a host that places it beside its own canvas (`renderSlideStage`) in a
 * custom shell instead of mounting the full `PptxViewer`.
 *
 * It is the same rail `PptxViewer` renders (`createThumbnailRail`), driven by
 * flat host options instead of the viewer store: the host owns the slides and
 * the active index, and every command (select, add, duplicate, delete, hide,
 * sections) is a callback. Thumbnails are drawn with the canvas's own inputs
 * (template elements, theme schemes, table styles and the deck-wide field
 * context, so `p:presentation/@firstSlideNum` numbering matches the canvas).
 *
 * The one difference is the thumbnail menu's "Layout" entry: in `PptxViewer`
 * it opens the built-in layout list, which needs the viewer's editor. Here it
 * makes the slide active and hands the request to `onOpenLayoutGallery`, so a
 * host can open its own Layout picker there.
 *
 * @module ui/slides-pane
 */
import {
	buildSlidesPaneFieldContext,
	deriveSlideFieldContext,
	openSlidesPaneLayoutGallery,
	visibleTemplateElements,
} from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	FieldSubstitutionContext,
	SlidesPaneLayoutGalleryRequest,
} from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxElement,
	PptxHeaderFooter,
	PptxSection,
	PptxSlide,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemeFormatScheme,
} from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { createTranslator } from '../i18n';
import { createDefaultRegistry, renderSlideStage } from '../render';
import type { ElementRendererRegistry } from '../render';
import { createThumbnailRail } from './thumbnails';
import type { ThumbnailRail, ThumbnailRailMenuDeps, ThumbnailSectionActions } from './thumbnails';

export type { SlidesPaneLayoutGalleryRequest };

/** Options of {@link createSlidesPane}. Every field can be changed later with `update`. */
export interface SlidesPaneOptions {
	/**
	 * The deck's slides in order, as the host stores them. When the host keeps
	 * layout and master elements apart (`templateElementsBySlideId`), pass them
	 * too so each thumbnail draws what the canvas draws; a deck straight from
	 * `loadPresentation` already has them merged in.
	 */
	slides: readonly PptxSlide[];
	/**
	 * Layout and master elements per slide id, drawn behind each slide's own
	 * elements (hidden ones skipped) exactly as `PptxViewer`'s canvas merges them.
	 */
	templateElementsBySlideId?: Readonly<Record<string, readonly PptxElement[]>>;
	/** Zero-based index of the slide shown on the canvas; its row gets `aria-current="true"`. */
	activeSlideIndex: number;
	/** The deck's slide size in CSS px (`LoadedPresentation.canvasSize`); thumbnails scale down from it. */
	canvasSize: CanvasSize;
	/** Resolved media URLs (`LoadedPresentation.mediaDataUrls`) for images and video posters. */
	mediaDataUrls?: ReadonlyMap<string, string>;
	/** The deck's sections; with any, the rail groups thumbnails under collapsible headers. */
	sections?: readonly PptxSection[];
	/**
	 * Whether the rail edits the deck: the thumbnail and section menus, Enter to
	 * insert after a thumbnail, and the "New slide" footer (shown only when
	 * `onAddSlide` is set too). Default `false`: read-only, no footer, no menus.
	 */
	canEdit?: boolean;
	/** Deck header/footer settings (`LoadedPresentation.headerFooter`), for date/footer fields. */
	headerFooter?: PptxHeaderFooter;
	/** Custom document properties (`LoadedPresentation.customProperties`), for property fields. */
	customProperties?: readonly PptxCustomProperty[];
	/**
	 * `p:presentation/@firstSlideNum` (`LoadedPresentation.firstSlideNumber`):
	 * the number a slide-number field shows on the first slide. Default 1.
	 */
	firstSlideNumber?: number;
	/** Locale for date/time fields; the canvas default when omitted. */
	locale?: string;
	/** Theme colours (`LoadedPresentation.colorScheme`), as passed to `renderSlideStage`. */
	colorScheme?: PptxThemeColorScheme;
	/** Theme fonts (`LoadedPresentation.fontScheme`), as passed to `renderSlideStage`. */
	fontScheme?: PptxThemeFontScheme;
	/** Theme format scheme (`LoadedPresentation.formatScheme`), as passed to `renderSlideStage`. */
	formatScheme?: PptxThemeFormatScheme;
	/** Table styles (`LoadedPresentation.tableStyleMap`) that table thumbnails resolve colours from. */
	tableStyleMap?: ParsedTableStyleMap;
	/** UI strings (labels, menus). Default: English (`createTranslator()`). */
	t?: Translator;
	/** Element renderers; pass the canvas's registry so custom renderers show in thumbnails too. */
	registry?: ElementRendererRegistry;
	/** A thumbnail was clicked: show that slide. Selecting never edits the deck. */
	onSelectSlide(index: number): void;
	/** The footer's "New slide" button. */
	onAddSlide?(): void;
	/** Insert a slide after `index` (the thumbnail menu's New Slide, and Enter on a thumbnail). */
	onAddSlideAfter?(index: number): void;
	/** Duplicate every selected slide (Ctrl/Shift click selects several). */
	onDuplicateSlides?(indexes: number[]): void;
	/** Delete every selected slide. */
	onDeleteSlides?(indexes: number[]): void;
	/** Toggle hidden on every selected slide. */
	onHideSlides?(indexes: number[]): void;
	/**
	 * The thumbnail menu's "Layout" entry. The rail first makes the slide
	 * active through `onSelectSlide`, then calls this with the slide and the
	 * click point, so the host can open its own Layout gallery there.
	 */
	onOpenLayoutGallery?(request: SlidesPaneLayoutGalleryRequest): void;
	/**
	 * Add a section named `name` starting at `afterSlideIndex` (the thumbnail
	 * menu's Add Section, and a section header's Add Section After).
	 */
	onAddSection?(name: string, afterSlideIndex: number): void;
	/** Rename a section (double-click its header, or the section menu's Rename). */
	onRenameSection?(sectionId: string, name: string): void;
	/** Delete a section; its slides stay in the deck. */
	onDeleteSection?(sectionId: string): void;
	/** Move a section, with its slides, one place up. */
	onMoveSectionUp?(sectionId: string): void;
	/** Move a section, with its slides, one place down. */
	onMoveSectionDown?(sectionId: string): void;
	/** Collapse or expand a section's thumbnails (click its header). */
	onToggleSection?(sectionId: string): void;
}

/** A mounted slide thumbnail rail; see {@link createSlidesPane}. */
export interface SlidesPane {
	/** The rail's root element (`<aside data-pptx-chrome="slides">`), appended to the container. */
	readonly el: HTMLElement;
	/**
	 * Apply changed options. Pass only what changed; omitted fields keep their
	 * value. Changing only `activeSlideIndex` or a callback does not redraw the
	 * thumbnails.
	 */
	update(options: Partial<SlidesPaneOptions>): void;
	/** Remove the rail from the container and release its thumbnails and menus. */
	destroy(): void;
}

/** Options that only change what a callback does or which row is current. */
const NON_RENDER_KEYS = new Set<keyof SlidesPaneOptions>([
	'activeSlideIndex',
	'onSelectSlide',
	'onAddSlide',
	'onAddSlideAfter',
	'onDuplicateSlides',
	'onDeleteSlides',
	'onHideSlides',
	'onOpenLayoutGallery',
	'onAddSection',
	'onRenameSection',
	'onDeleteSection',
	'onMoveSectionUp',
	'onMoveSectionDown',
	'onToggleSection',
]);

/**
 * Mount the slide thumbnail rail into `container` (appended as its last
 * child) and return a handle to update or remove it.
 *
 * ```ts
 * const pane = createSlidesPane(railHost, {
 * 	slides: loaded.slides,
 * 	activeSlideIndex: 0,
 * 	canvasSize: loaded.canvasSize,
 * 	mediaDataUrls: loaded.mediaDataUrls,
 * 	firstSlideNumber: loaded.firstSlideNumber,
 * 	onSelectSlide: (index) => showSlide(index),
 * });
 * pane.update({ activeSlideIndex: 2 });
 * ```
 *
 * The rail is styled by the viewer stylesheet (`getViewerCss()` or
 * `pptx-vanilla-viewer/styles.css`) scoped under a `.pptxv` ancestor.
 */
export function createSlidesPane(container: HTMLElement, initial: SlidesPaneOptions): SlidesPane {
	const doc = container.ownerDocument;
	let options: SlidesPaneOptions = { ...initial };
	let destroyed = false;
	let defaultTranslator: Translator | undefined;
	let defaultRegistry: ElementRendererRegistry | undefined;
	const translator = (): Translator => options.t ?? (defaultTranslator ??= createTranslator());
	const registry = (): ElementRendererRegistry =>
		options.registry ?? (defaultRegistry ??= createDefaultRegistry());

	// Deck-wide field context, rebuilt with each render pass; every thumbnail
	// re-points the slide number and title at its own slide.
	let deckFieldContext: FieldSubstitutionContext = buildSlidesPaneFieldContext({});

	const select = (index: number): void => options.onSelectSlide(index);

	const renderStage = (slide: PptxSlide, scale: number): HTMLElement => {
		const template = visibleTemplateElements(
			slide,
			(options.templateElementsBySlideId?.[slide.id] ?? []) as PptxElement[],
		);
		return renderSlideStage({
			document: doc,
			slide: template.length ? { ...slide, elements: [...template, ...slide.elements] } : slide,
			canvasSize: options.canvasSize,
			mediaDataUrls: options.mediaDataUrls ?? new Map(),
			colorScheme: options.colorScheme,
			fontScheme: options.fontScheme,
			formatScheme: options.formatScheme,
			tableStyleMap: options.tableStyleMap,
			fieldContext: deriveSlideFieldContext(deckFieldContext, slide),
			registry: registry(),
			t: translator(),
			scale,
			slides: options.slides,
			currentSlideIndex: options.activeSlideIndex,
		});
	};

	const menuDeps: ThumbnailRailMenuDeps = {
		addSlideAfter: (index) => options.onAddSlideAfter?.(index),
		duplicateSlides: (indexes) => options.onDuplicateSlides?.(indexes),
		deleteSlides: (indexes) => options.onDeleteSlides?.(indexes),
		toggleHideSlides: (indexes) => options.onHideSlides?.(indexes),
		openLayout: (slideIndex, x, y) =>
			openSlidesPaneLayoutGallery({ slideIndex, x, y }, select, options.onOpenLayoutGallery),
		addSection: (name, afterSlideIndex) => options.onAddSection?.(name, afterSlideIndex),
	};

	const sectionActions: ThumbnailSectionActions = {
		toggle: (sectionId) => options.onToggleSection?.(sectionId),
		rename: (sectionId, name) => options.onRenameSection?.(sectionId, name),
		delete: (sectionId) => options.onDeleteSection?.(sectionId),
		move: (sectionId, direction) =>
			direction === 'up'
				? options.onMoveSectionUp?.(sectionId)
				: options.onMoveSectionDown?.(sectionId),
		addAfter: (slideIndex) =>
			options.onAddSection?.(translator()('pptx.sections.defaultName'), slideIndex),
	};

	// The menus and the Enter shortcut are wired when the rail is built, so a
	// change of `canEdit` or translator builds a fresh rail in the same place.
	let rail: ThumbnailRail | null = null;
	let builtFor: { canEdit: boolean; t: Translator } | null = null;
	const ensureRail = (): ThumbnailRail => {
		const canEdit = options.canEdit === true;
		const t = translator();
		if (rail && builtFor?.canEdit === canEdit && builtFor.t === t) {
			return rail;
		}
		const next = createThumbnailRail(
			doc,
			t,
			select,
			() => options.onAddSlide?.(),
			canEdit ? menuDeps : undefined,
		);
		if (rail) {
			rail.destroy();
			rail.el.replaceWith(next.el);
		} else {
			container.appendChild(next.el);
		}
		rail = next;
		builtFor = { canEdit, t };
		return next;
	};

	const render = (): void => {
		const target = ensureRail();
		deckFieldContext = buildSlidesPaneFieldContext({
			headerFooter: options.headerFooter,
			customProperties: options.customProperties,
			firstSlideNumber: options.firstSlideNumber,
			locale: options.locale,
		});
		const canEdit = options.canEdit === true;
		target.setAddSlideVisible(canEdit && Boolean(options.onAddSlide));
		target.setActive(options.activeSlideIndex);
		target.render(
			[...options.slides],
			options.canvasSize,
			renderStage,
			options.sections ?? [],
			canEdit ? sectionActions : undefined,
		);
	};

	render();
	let mounted = rail!.el;

	return {
		get el() {
			mounted = rail?.el ?? mounted;
			return mounted;
		},
		update(patch) {
			if (destroyed) {
				return;
			}
			const previous = options;
			options = { ...options, ...patch };
			const changed = (Object.keys(patch) as (keyof SlidesPaneOptions)[]).filter(
				(key) => previous[key] !== options[key],
			);
			if (changed.some((key) => !NON_RENDER_KEYS.has(key))) {
				render();
				return;
			}
			if (changed.includes('onAddSlide')) {
				rail?.setAddSlideVisible(options.canEdit === true && Boolean(options.onAddSlide));
			}
			if (changed.includes('activeSlideIndex')) {
				rail?.setActive(options.activeSlideIndex);
			}
		},
		destroy() {
			if (destroyed) {
				return;
			}
			destroyed = true;
			rail?.destroy();
			rail?.el.remove();
			rail = null;
		},
	};
}
