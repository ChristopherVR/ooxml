import { BEVEL_PRESETS, MATERIAL_PRESETS } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n/translator';

/**
 * text-effects-labels.ts: display labels for the two preset selects in
 * {@link TextEffectsSection} (3D material, 3D bevel). The warp picker is the
 * shared Text Warp gallery, which names its own tiles.
 *
 * Every binding translates the same shared catalogues through its own
 * translator (`labelKey`), so the wording cannot drift between the bindings.
 *
 * These helpers are LOOKUPS only. They never decide which values a select
 * offers, so spelling a token out cannot add or drop an option (which would
 * move the control out of parity with React).
 */

/** Index a shared preset catalogue's translation keys by its wire value. */
function keyIndex(
	presets: ReadonlyArray<{ value: string; labelKey: string }>,
): ReadonlyMap<string, string> {
	return new Map(presets.map((preset) => [preset.value, preset.labelKey]));
}

const MATERIAL_KEYS = keyIndex(MATERIAL_PRESETS);
const BEVEL_KEYS = keyIndex(BEVEL_PRESETS);

/**
 * An unmapped token falls back to itself rather than to an empty string: a
 * deck may carry a preset newer than the catalogue, and blanking the option
 * would make the select look broken rather than merely untranslated.
 */
function lookup(keys: ReadonlyMap<string, string>, value: string, t: Translator): string {
	const key = keys.get(value);
	return key === undefined ? value : t(key);
}

/** Display label for a `Text3DStyle.presetMaterial` value. */
export function materialPresetLabel(value: string, t: Translator): string {
	return lookup(MATERIAL_KEYS, value, t);
}

/** Display label for a `Text3DStyle.bevelTopType` value. */
export function bevelPresetLabel(value: string, t: Translator): string {
	return lookup(BEVEL_KEYS, value, t);
}
