import type { PptxCustomProperty, PptxHeaderFooter } from 'ooxml-core/pptx';

import { buildFieldSubstitutionContext } from './field-context';
import type { FieldSubstitutionContext } from './text-field-substitution';

/**
 * Framework-free wiring for the slides pane (the slide thumbnail rail) when a
 * host composes it beside its own canvas instead of using the full viewer.
 *
 * Every binding exposes its own rail component; this module holds what they
 * share, so each binding only maps it onto its template: the deck-wide field
 * and table style contexts the thumbnails need to draw what the canvas draws,
 * the rehearsal timings shown under each thumbnail, the bulk slide commands
 * the rail's context menu issues, and the hand-off of the "Layout" entry to a
 * host that renders its own Layout gallery.
 */

/** Deck-level inputs of the rail's thumbnail field context. */
export interface SlidesPaneFieldContextSources {
	headerFooter?: PptxHeaderFooter;
	customProperties?: readonly PptxCustomProperty[];
	/** `p:presentation@firstSlideNum`; numbering starts at 1 when undefined. */
	firstSlideNumber?: number | undefined;
	locale?: string;
}

/**
 * The deck-wide field context for the rail's thumbnails. Built with the same
 * helper as the canvas, without a slide: each thumbnail re-points the slide
 * number and title at its own slide, so one context serves every thumbnail and
 * stays stable while the active slide changes.
 */
export function buildSlidesPaneFieldContext(
	sources: SlidesPaneFieldContextSources,
): FieldSubstitutionContext {
	return buildFieldSubstitutionContext({
		headerFooter: sources.headerFooter,
		customProperties: sources.customProperties,
		firstSlideNumber: sources.firstSlideNumber,
		...(sources.locale === undefined ? {} : { locale: sources.locale }),
	});
}

/**
 * The rail's table style context: the theme and table style map the canvas
 * resolves table colours from, or `undefined` when the deck has neither (the
 * canvas's "no context" fallback). Generic so each binding keeps its own
 * theme and style map types.
 */
export function buildSlidesPaneTableStyleContext<TTheme, TStyleMap>(
	theme: TTheme | undefined,
	tableStyleMap: TStyleMap | undefined,
): { theme: TTheme | undefined; tableStyleMap: TStyleMap | undefined } | undefined {
	return theme || tableStyleMap ? { theme, tableStyleMap } : undefined;
}

/**
 * Rehearsal timings (slide index to seconds) for the rail, or `undefined` when
 * none were recorded, so the rail shows no timing row on an unrehearsed deck.
 */
export function slidesPaneRehearsalTimings(
	recordedTimings: Readonly<Record<number, number>> | undefined,
): Record<number, number> | undefined {
	return recordedTimings && Object.keys(recordedTimings).length > 0
		? { ...recordedTimings }
		: undefined;
}

/** Where the rail asked for the Layout gallery: the slide, and the click point. */
export interface SlidesPaneLayoutGalleryRequest {
	/** Zero-based index of the slide whose layout is being changed (now the active slide). */
	slideIndex: number;
	/** Viewport x of the context-menu click, to anchor the gallery. */
	x: number;
	/** Viewport y of the context-menu click, to anchor the gallery. */
	y: number;
}

/**
 * The rail context menu's "Layout" entry for a host that renders its own
 * Layout gallery: makes the slide active, then hands the request to the host.
 * Without a host handler the slide still becomes active, so the entry is never
 * a silent no-op.
 */
export function openSlidesPaneLayoutGallery(
	request: SlidesPaneLayoutGalleryRequest,
	setActiveSlide: (index: number) => void,
	onOpenLayoutGallery: ((request: SlidesPaneLayoutGalleryRequest) => void) | undefined,
): void {
	setActiveSlide(request.slideIndex);
	onOpenLayoutGallery?.(request);
}

/** Single-slide operations a binding already has, lifted to the rail's multi-selection. */
export interface SlidesPaneSingleSlideOps {
	duplicateSlide: (index: number) => void;
	deleteSlide: (index: number) => void;
	toggleSlideHidden: (index: number) => void;
}

/** The rail's multi-selection commands. */
export interface SlidesPaneBulkSlideOps {
	duplicateSlides: (indexes: readonly number[]) => void;
	deleteSlides: (indexes: readonly number[]) => void;
	toggleHideSlides: (indexes: readonly number[]) => void;
}

/**
 * Bulk rail commands over single-slide operations. Deletion runs from the
 * highest index down so earlier deletions do not shift the indexes still to
 * delete; duplication runs from the highest index down so each copy lands
 * after its own source. Indexes are de-duplicated first.
 */
export function createSlidesPaneBulkSlideOps(
	ops: SlidesPaneSingleSlideOps,
): SlidesPaneBulkSlideOps {
	const descending = (indexes: readonly number[]): number[] =>
		[...new Set(indexes)].sort((a, b) => b - a);
	return {
		duplicateSlides: (indexes) => {
			for (const index of descending(indexes)) {
				ops.duplicateSlide(index);
			}
		},
		deleteSlides: (indexes) => {
			for (const index of descending(indexes)) {
				ops.deleteSlide(index);
			}
		},
		toggleHideSlides: (indexes) => {
			for (const index of new Set(indexes)) {
				ops.toggleSlideHidden(index);
			}
		},
	};
}
