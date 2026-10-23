/**
 * smart-art-presets.ts: the SmartArt insert-gallery catalogue as PowerPoint names it. The
 * catalogue itself is product-neutral (`DIAGRAM_INSERT_PRESETS` / `DIAGRAM_INSERT_CATEGORIES` in
 * `ooxml-core/diagram`); this module adds the pptx translation keys and keeps the historical
 * `PRESETS` / `CATEGORIES` import surface every binding and test uses.
 */

import { DIAGRAM_INSERT_CATEGORIES, DIAGRAM_INSERT_PRESETS } from 'ooxml-core/diagram';
import type { DiagramInsertCategory } from 'ooxml-core/diagram';
import type { SmartArtLayout } from 'ooxml-core/pptx';

/** The five categories surfaced in the dialog's left sidebar. */
export type SmartArtCategory = DiagramInsertCategory;

/** A single entry in the SmartArt insert gallery. */
export interface SmartArtPreset {
	/** Core layout kind this preset inserts. */
	layout: SmartArtLayout;
	/** Display label (English fallback). */
	label: string;
	/** Translation key resolving to the label. */
	labelKey: string;
	/** Sidebar category. */
	category: SmartArtCategory;
	/** Default node texts to seed the new SmartArt with. */
	defaultItems: string[];
}

/** Dictionary key of a preset's name. */
export const smartArtPresetLabelKey = (layout: string): string => `pptx.smartart.preset.${layout}`;
/** Dictionary key of an insert category's name. */
export const smartArtCategoryLabelKey = (category: string): string =>
	`pptx.smartart.category.${category}`;
/** Dictionary key of an insert category's description. */
export const smartArtCategoryDescriptionKey = (category: string): string =>
	`pptx.insertSmartArt.description.${category}`;

export const PRESETS: SmartArtPreset[] = DIAGRAM_INSERT_PRESETS.map((preset) => ({
	layout: preset.layout,
	label: preset.label,
	labelKey: smartArtPresetLabelKey(preset.layout),
	category: preset.category,
	defaultItems: [...preset.defaultItems],
}));

/** Sidebar tabs, in display order. */
export const CATEGORIES: Array<{ id: SmartArtCategory; label: string; labelKey: string }> =
	DIAGRAM_INSERT_CATEGORIES.map((category) => ({
		id: category.id,
		label: category.label,
		labelKey: smartArtCategoryLabelKey(category.id),
	}));
