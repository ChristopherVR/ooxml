import { PATTERN_PRESET_OPTIONS as SHARED_PATTERN_PRESET_OPTIONS } from 'ooxml-ui/pptx';

// ---------------------------------------------------------------------------
// Option arrays for fill / stroke properties
//
// NOTE: `label` keeps the English fallback text (existing consumers still
// render `option.label` directly). Each option also carries an `i18nKey`
// pointing at the shared i18n dictionary, matching the `{ value, i18nKey }`
// convention already used elsewhere in this codebase, so a render site can
// switch to `t(option.i18nKey)` without a data-shape change.
//
// PATTERN_PRESET_OPTIONS used to be a private, byte-identical retype of the
// shared `render/fill-pattern-label-keys.ts` catalogue; it is the one shared
// copy now. The line options are drawn from the shared Line descriptors
// (`SHAPE_LINE_SECTIONS`).
// ---------------------------------------------------------------------------

/** Shape is `{ value, labelKey }` (no English `label`); render via `t(opt.labelKey)`. */
export const PATTERN_PRESET_OPTIONS = SHARED_PATTERN_PRESET_OPTIONS;

export const FILL_MODE_OPTIONS = [
	{ value: 'solid', label: 'Solid', i18nKey: 'pptx.fill.solid' },
	{ value: 'gradient', label: 'Gradient', i18nKey: 'pptx.fill.gradient' },
	{ value: 'pattern', label: 'Pattern', i18nKey: 'pptx.table.fillPattern' },
	{ value: 'image', label: 'Image', i18nKey: 'pptx.inspector.image' },
	{ value: 'none', label: 'None', i18nKey: 'pptx.fill.none' },
];

export const GRADIENT_TYPE_OPTIONS = [
	{ value: 'linear', label: 'Linear', i18nKey: 'pptx.gradient.linear' },
	{ value: 'radial', label: 'Radial', i18nKey: 'pptx.gradient.radial' },
];

export const IMAGE_MODE_OPTIONS = [
	{ value: 'stretch', label: 'Stretch', i18nKey: 'pptx.image.stretch' },
	{ value: 'tile', label: 'Tile', i18nKey: 'pptx.image.tile' },
];
