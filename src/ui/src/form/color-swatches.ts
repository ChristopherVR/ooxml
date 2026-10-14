/**
 * color-swatches.ts: the canonical "Office Standard Colors" swatch catalogue
 * shared by every binding's colour pickers (font colour, highlight colour,
 * and any future fill/line colour picker that wants the same 10-swatch grid).
 *
 * Mirrors the "Standard Colors" row PowerPoint itself ships. Before this
 * module existed, the VanillaJS (`swatch-picker.ts`) and Svelte
 * (`SwatchColorPicker.svelte`) bindings each independently hardcoded an
 * identical 10-colour list; this module is the single source of truth going
 * forward so no binding drifts from the others.
 *
 * Each swatch carries the dictionary key of its name (`labelKey`) beside the
 * English `label`; {@link officeColorSwatchLabel} resolves it through the
 * binding's translator so a tooltip follows the viewer's locale.
 *
 * @module render/color-swatches
 */
import { translateOr, type Translate } from 'ooxml-core/i18n';

/** One entry in the Office standard colour swatch grid. */
export interface OfficeColorSwatch {
	/** Swatch colour as a lower-case `#rrggbb` hex string. */
	hex: string;
	/** English fallback name for the swatch (used for tooltips/aria-labels). */
	label: string;
	/** Dictionary key of the name (`pptx.colorPicker.standard.*`). */
	labelKey: string;
}

/**
 * The canonical 10-swatch "Office Standard Colors" catalogue, in PowerPoint's
 * on-screen order. Consumed by the font-colour and highlight-colour pickers
 * (and any other picker that wants the same standard set) across every
 * binding.
 */
export const OFFICE_COLOR_SWATCHES: readonly OfficeColorSwatch[] = [
	{ hex: '#000000', label: 'Black', labelKey: 'pptx.colorPicker.standard.black' },
	{ hex: '#ffffff', label: 'White', labelKey: 'pptx.colorPicker.standard.white' },
	{ hex: '#ff0000', label: 'Red', labelKey: 'pptx.colorPicker.standard.red' },
	{ hex: '#00aa00', label: 'Green', labelKey: 'pptx.colorPicker.standard.green' },
	{ hex: '#0000ff', label: 'Blue', labelKey: 'pptx.colorPicker.standard.blue' },
	{ hex: '#ff8800', label: 'Orange', labelKey: 'pptx.colorPicker.standard.orange' },
	{ hex: '#8800cc', label: 'Purple', labelKey: 'pptx.colorPicker.standard.purple' },
	{ hex: '#00cccc', label: 'Cyan', labelKey: 'pptx.colorPicker.standard.cyan' },
	{ hex: '#ff69b4', label: 'Pink', labelKey: 'pptx.colorPicker.standard.pink' },
	{ hex: '#808080', label: 'Gray', labelKey: 'pptx.colorPicker.standard.gray' },
];

/**
 * The same catalogue flattened to bare `#rrggbb` hex strings, in the same
 * order as {@link OFFICE_COLOR_SWATCHES}. A drop-in replacement for the
 * `readonly string[]` swatch lists the vanilla and Svelte bindings each
 * hardcoded locally before this module existed.
 */
export const OFFICE_COLOR_SWATCH_HEXES: readonly string[] = OFFICE_COLOR_SWATCHES.map(
	(swatch) => swatch.hex,
);

/** A standard swatch's name in the viewer's locale (English without a translator or a key). */
export function officeColorSwatchLabel(swatch: OfficeColorSwatch, translate?: Translate): string {
	return translateOr(translate, swatch.labelKey, swatch.label);
}
