/**
 * slides-panel-contexts.ts: the field and table style contexts the slide
 * thumbnail rail (`SlidesPanelComponent`) hands its thumbnails, so each
 * miniature draws what the canvas draws.
 *
 * Built with the shared slides-pane helpers from `ooxml-ui/pptx`. Each source
 * is the rail's own input when the host sets it, otherwise the enclosing deck
 * loader's value (the same signal the canvas's `FieldContextService` and table
 * renderers read), so the rail matches the canvas both inside
 * `PowerPointViewerComponent` and in a custom shell that feeds it from inputs.
 *
 * @module angular-viewer/slides-panel-contexts
 */
import { computed } from '@angular/core';
import type { Signal } from '@angular/core';
import {
	buildSlidesPaneFieldContext,
	buildSlidesPaneTableStyleContext,
	deriveSlideFieldContext,
} from 'ooxml-ui/pptx';
import type { FieldSubstitutionContext, TableStyleContext } from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxHeaderFooter,
	PptxSlide,
	PptxTheme,
} from 'pptx-viewer-core';

import type { LoadContentService } from './load-content.service';
import { tableStyleContextFor } from './table-renderer-helpers';

/** The rail inputs a thumbnail context is built from (`undefined` falls back to the loader). */
export interface SlidesPanelContextInputs {
	headerFooter: Signal<PptxHeaderFooter | undefined>;
	customProperties: Signal<readonly PptxCustomProperty[] | undefined>;
	firstSlideNumber: Signal<number | undefined>;
	locale: Signal<string | undefined>;
	theme: Signal<PptxTheme | undefined>;
	tableStyleMap: Signal<ParsedTableStyleMap | undefined>;
}

/**
 * Field context provider for the thumbnails, standing in for the viewer's
 * `FieldContextService` (the slide canvas only calls {@link forSlide}). The
 * deck-wide context is shared; each thumbnail re-points the slide number and
 * title at its own slide.
 */
export class SlidesPanelFieldContext {
	constructor(
		/** Deck-wide fields: header/footer, custom properties, first slide number. */
		readonly deckContext: Signal<FieldSubstitutionContext>,
	) {}

	/** The deck-wide context re-pointed at `slide`'s number and title. */
	forSlide(slide: PptxSlide | undefined): FieldSubstitutionContext {
		return deriveSlideFieldContext(this.deckContext(), slide) ?? this.deckContext();
	}
}

/** The rail's thumbnail contexts. */
export interface SlidesPanelContexts {
	fields: SlidesPanelFieldContext;
	tableStyle: Signal<TableStyleContext | undefined>;
}

/** Build the thumbnail contexts from the rail inputs, falling back to `loader` per source. */
export function createSlidesPanelContexts(
	inputs: SlidesPanelContextInputs,
	loader: LoadContentService | null,
): SlidesPanelContexts {
	const deckContext = computed(() =>
		buildSlidesPaneFieldContext({
			headerFooter: inputs.headerFooter() ?? loader?.headerFooter(),
			customProperties: inputs.customProperties() ?? loader?.customProperties(),
			firstSlideNumber: inputs.firstSlideNumber() ?? loader?.firstSlideNumber(),
			locale: inputs.locale(),
		}),
	);
	const tableStyle = computed(() => {
		const context = buildSlidesPaneTableStyleContext(
			inputs.theme() ?? loader?.theme(),
			inputs.tableStyleMap() ?? loader?.tableStyleMap(),
		);
		return context ? tableStyleContextFor(context.tableStyleMap, context.theme) : undefined;
	});
	return { fields: new SlidesPanelFieldContext(deckContext), tableStyle };
}
