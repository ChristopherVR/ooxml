import { buildSlidesPaneFieldContext, buildSlidesPaneTableStyleContext } from 'ooxml-ui/pptx';
import type { FieldSubstitutionContext } from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxCustomProperty,
	PptxHeaderFooter,
	PptxTheme,
} from 'pptx-viewer-core';
import { provide } from 'vue';

import { FieldContextKey } from './field-context';
import { TableThemeKey, tableThemeContextFor } from './table-theme';
import type { TableThemeContext } from './table-theme';

/**
 * Deck-level state the slide thumbnail rail needs so its thumbnails draw what
 * the canvas draws. Every field is optional: an omitted field reads as "the
 * deck has none", exactly as the canvas treats it.
 */
export interface SlidesPaneDeckSources {
	/** Deck header/footer settings (date, footer and header field text). */
	headerFooter?: PptxHeaderFooter | undefined;
	/** Custom document properties, for `docproperty` field runs. */
	customProperties?: readonly PptxCustomProperty[] | undefined;
	/** `p:presentation/@firstSlideNum`; numbering starts at 1 when undefined. */
	firstSlideNumber?: number | undefined;
	/** Locale for date and time fields (browser default when unset). */
	locale?: string | undefined;
	/** Presentation theme, for table colour, font and format scheme lookups. */
	theme?: PptxTheme | undefined;
	/** Parsed `ppt/tableStyles.xml`, for table style lookups. */
	tableStyleMap?: ParsedTableStyleMap | undefined;
}

/** The deck-wide field context of the rail's thumbnails (no slide number or title). */
export function slidesPaneFieldContext(sources: SlidesPaneDeckSources): FieldSubstitutionContext {
	return buildSlidesPaneFieldContext({
		headerFooter: sources.headerFooter,
		customProperties: sources.customProperties,
		firstSlideNumber: sources.firstSlideNumber,
		...(sources.locale === undefined ? {} : { locale: sources.locale }),
	});
}

/** The rail's table theme context, or `undefined` when the deck has no theme or table styles. */
export function slidesPaneTableTheme(
	sources: SlidesPaneDeckSources,
): TableThemeContext | undefined {
	const context = buildSlidesPaneTableStyleContext(sources.theme, sources.tableStyleMap);
	return context ? tableThemeContextFor(context.theme, context.tableStyleMap) : undefined;
}

/**
 * Provide the rail's field and table theme contexts to its subtree (call from
 * `setup`). The thumbnails' `SlideStage` re-points the field context at the
 * slide each one paints, so one deck-wide value serves every thumbnail. Both
 * are built from the host's props only, so the rail renders the same inside
 * `PowerPointViewer` and beside a host's own canvas.
 */
export function provideSlidesPaneContexts(sources: () => SlidesPaneDeckSources): void {
	provide(FieldContextKey, () => slidesPaneFieldContext(sources()));
	provide(TableThemeKey, () => slidesPaneTableTheme(sources()));
}
