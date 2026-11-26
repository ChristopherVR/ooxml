/**
 * `animation-direction-options` - which directions the Animation pane offers
 * for a preset: only the ones PowerPoint itself has a variant (a
 * `presetSubtype`) for, so a picked direction always saves as a subtype and
 * behaviour tree PowerPoint recognises.
 *
 * Ground truth: every `MsoAnimEffect` was driven through COM with every
 * `MsoAnimDirection` it accepts and the saved subtype read back (core's
 * `animation-behavior-captures.json`). Fly and Crawl accept all eight
 * directions (edges and corners); Wipe, Peek and Stretch accept the four
 * edges; Float In has no direction at all (Float Up / Float Down are
 * separate presets), so it no longer shows a picker that saved nothing.
 *
 * @module render/animation-direction-options
 */
import {
	getCapturedPreset,
	resolveDirectionSubtype,
	resolveOoxmlPresetMapping,
	type PptxAnimationDirection,
	type PptxElementAnimation,
} from 'ooxml-core/pptx';

/** Every direction value, edges first then corners. */
export const DIRECTION_VALUES: readonly PptxAnimationDirection[] = [
	'fromTop',
	'fromBottom',
	'fromLeft',
	'fromRight',
	'fromTopLeft',
	'fromTopRight',
	'fromBottomLeft',
	'fromBottomRight',
];

/**
 * Direction -> i18n dictionary key, for every picker in every binding, so the
 * bindings spell a direction the same way without each keeping its own map.
 */
export const ANIMATION_DIRECTION_LABEL_KEYS: Readonly<Record<PptxAnimationDirection, string>> = {
	fromTop: 'pptx.animation.direction.fromTop',
	fromBottom: 'pptx.animation.direction.fromBottom',
	fromLeft: 'pptx.animation.direction.fromLeft',
	fromRight: 'pptx.animation.direction.fromRight',
	fromTopLeft: 'pptx.animation.direction.fromTopLeft',
	fromTopRight: 'pptx.animation.direction.fromTopRight',
	fromBottomLeft: 'pptx.animation.direction.fromBottomLeft',
	fromBottomRight: 'pptx.animation.direction.fromBottomRight',
	horizontal: 'pptx.animation.direction.horizontal',
	vertical: 'pptx.animation.direction.vertical',
	horizontalIn: 'pptx.animation.direction.horizontalIn',
	horizontalOut: 'pptx.animation.direction.horizontalOut',
	verticalIn: 'pptx.animation.direction.verticalIn',
	verticalOut: 'pptx.animation.direction.verticalOut',
	in: 'pptx.animation.direction.in',
	out: 'pptx.animation.direction.out',
	objectCenter: 'pptx.animation.direction.objectCenter',
	slideCenter: 'pptx.animation.direction.slideCenter',
	inSlightly: 'pptx.animation.direction.inSlightly',
	outSlightly: 'pptx.animation.direction.outSlightly',
	inFromScreenCenter: 'pptx.animation.direction.inFromScreenCenter',
	outFromScreenCenter: 'pptx.animation.direction.outFromScreenCenter',
	outFromScreenBottom: 'pptx.animation.direction.outFromScreenBottom',
	inToScreenBottom: 'pptx.animation.direction.inToScreenBottom',
};

const EDGES: readonly PptxAnimationDirection[] = DIRECTION_VALUES.slice(0, 4);

const PRESET_DIRECTIONS: Readonly<Record<string, readonly PptxAnimationDirection[]>> = {
	flyIn: DIRECTION_VALUES,
	flyOut: DIRECTION_VALUES,
	crawlIn: DIRECTION_VALUES,
	crawlOut: DIRECTION_VALUES,
	wipeIn: EDGES,
	wipeOut: EDGES,
	peekIn: EDGES,
	peekOut: EDGES,
	stretchIn: EDGES,
};

/** Presets that expose the direction picker. */
export const DIRECTIONAL_PRESETS: ReadonlySet<string> = new Set(Object.keys(PRESET_DIRECTIONS));

/** The directions `preset` offers, empty when it has none. */
export function directionValuesForPreset(
	preset: string | undefined,
): readonly PptxAnimationDirection[] {
	if (!preset) return [];
	const edges = PRESET_DIRECTIONS[preset];
	if (edges) return edges;
	const mapping = resolveOoxmlPresetMapping(preset);
	if (!mapping || (mapping.presetClass !== 'entr' && mapping.presetClass !== 'exit')) return [];
	switch (mapping.presetId) {
		case 2:
		case 7:
			return DIRECTION_VALUES;
		case 12:
		case 17:
		case 22:
			return EDGES;
		case 3:
			return ['horizontal', 'vertical'];
		case 4:
		case 6:
		case 8:
		case 13:
			return ['in', 'out'];
		case 16:
			return ['verticalIn', 'horizontalIn', 'verticalOut', 'horizontalOut'];
		case 23:
			return mapping.presetClass === 'entr'
				? ['in', 'inSlightly', 'out', 'outSlightly', 'inFromScreenCenter', 'outFromScreenBottom']
				: ['in', 'inSlightly', 'out', 'outSlightly', 'outFromScreenCenter', 'inToScreenBottom'];
		case 53:
			return ['objectCenter', 'slideCenter'];
		default:
			return [];
	}
}

/** Read a native direction by comparing its subtype to the choices the writer emits. */
export function directionForAnimationSubtype(
	preset: string,
	subtype: number | undefined,
): PptxAnimationDirection | undefined {
	const mapping = resolveOoxmlPresetMapping(
		preset,
		(cls, id) => getCapturedPreset(cls, id)?.defaultSubtype,
	);
	return mapping && subtype !== undefined
		? directionValuesForPreset(preset).find(
				(value) => resolveDirectionSubtype(mapping, value) === subtype,
			)
		: undefined;
}

/**
 * The direction a picker shows for an entry: its own, else PowerPoint's
 * default for every directional preset here ("From Bottom", subtype 4;
 * Stretch's default "Across" has no picker value, so its first edge).
 */
export function effectiveDirection(
	entry: Pick<PptxElementAnimation, 'direction'> | undefined,
	values: readonly PptxAnimationDirection[],
): PptxAnimationDirection | undefined {
	if (entry?.direction && values.includes(entry.direction)) {
		return entry.direction;
	}
	return values.includes('fromBottom') ? 'fromBottom' : values[0];
}

/**
 * The directions offered for an element's animation entry: its entrance's,
 * else its exit's (the pane edits one `direction` per entry). Empty when the
 * element has no directional preset.
 */
export function directionValuesFor(
	slideAnimations: readonly PptxElementAnimation[],
	elementId: string,
): readonly PptxAnimationDirection[] {
	const entry = slideAnimations.find((a) => a.elementId === elementId);
	if (!entry) {
		return [];
	}
	const entrance = directionValuesForPreset(entry.entrance);
	return entrance.length > 0 ? entrance : directionValuesForPreset(entry.exit);
}
