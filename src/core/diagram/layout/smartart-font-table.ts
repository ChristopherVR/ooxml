/**
 * Font advance-table lookup shared by the layout engine and the interpreters.
 * Moved from `pptx/core/utils/smartart-layout-item-font-size.ts`, which
 * re-exports it.
 */

import {
	DEFAULT_FONT_ADVANCE_TABLE,
	FONT_ADVANCE_TABLES,
} from '../../text/font-metrics/font-advance-widths.generated';
import type { FontAdvanceTable } from '../../text/font-metrics/font-advance-widths.generated';

/**
 * Resolve the {@link FontAdvanceTable} for `fontName`, defaulting to Calibri
 * (the Office theme's default minor font, and what SmartArt renders in
 * absent an explicit per-node font override) and falling back to the
 * cross-font average for a font this table has not measured.
 */
export function resolveFontTable(fontName: string | undefined): FontAdvanceTable {
	return FONT_ADVANCE_TABLES[fontName ?? 'Calibri'] ?? DEFAULT_FONT_ADVANCE_TABLE;
}
