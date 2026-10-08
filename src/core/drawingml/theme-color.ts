// Resolution of scheme colour names (`accent1`, `tx1`, `bg2`...) against a theme colour scheme
// through a colour map, and the bridge to the DrawingML colour resolver.
import type { DrawingColorTheme } from './drawing-color';
import { isThemeColorSlot } from './theme';
import {
	DEFAULT_THEME_COLOR_MAP,
	type ThemeColorMap,
	type ThemeColorScheme,
	type ThemeColorSlot,
	type ThemeLogicalColor,
} from './theme-model';
import type { DrawingColor } from './types';

function isLogical(name: string): name is ThemeLogicalColor {
	return Object.hasOwn(DEFAULT_THEME_COLOR_MAP, name);
}

/**
 * The scheme slot a scheme colour name uses. `bg1`/`tx1`/`bg2`/`tx2` go through `colorMap`,
 * falling back to the conventional mapping; slot names (`dk1`, `accent3`) stand for themselves
 * unless the map reassigns them; anything else (`phClr`) has no slot.
 */
export function themeColorSlotFor(
	name: string,
	colorMap: ThemeColorMap = {},
): ThemeColorSlot | undefined {
	if (isLogical(name)) return colorMap[name] ?? DEFAULT_THEME_COLOR_MAP[name];
	if (!isThemeColorSlot(name)) return undefined;
	if (name === 'dk1' || name === 'lt1' || name === 'dk2' || name === 'lt2') return name;
	return colorMap[name] ?? name;
}

/** The colour of a scheme name, through the colour map; `undefined` when the scheme lacks it. */
export function resolveSchemeColor(
	scheme: ThemeColorScheme,
	name: string,
	colorMap?: ThemeColorMap,
): DrawingColor | undefined {
	const slot = themeColorSlotFor(name, colorMap);
	return slot ? scheme.colors[slot] : undefined;
}

/**
 * The `RRGGBB` (upper case, no `#`) a scheme slot colour stands for: an `srgbClr` value or a
 * `sysClr`'s `lastClr`. Other kinds are not plain values in a scheme and give `undefined`.
 */
export function themeSlotHex(color: DrawingColor | undefined): string | undefined {
	if (!color) return undefined;
	const value =
		color.kind === 'srgb' ? color.value : color.kind === 'system' ? color.fallback : undefined;
	return value ? value.toUpperCase() : undefined;
}

/** A `DrawingColorTheme` (`#RRGGBB` per scheme name) over a colour scheme and optional colour map. */
export function themeDrawingColorTheme(
	scheme: ThemeColorScheme | undefined,
	colorMap?: ThemeColorMap,
): DrawingColorTheme {
	return {
		scheme(name) {
			const hex = scheme && themeSlotHex(resolveSchemeColor(scheme, name, colorMap));
			return hex ? `#${hex.replace(/^#/, '')}` : undefined;
		},
	};
}
