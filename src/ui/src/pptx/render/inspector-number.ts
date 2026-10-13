/**
 * Display rounding for the inspector's numeric inputs.
 *
 * Many model values are converted from EMU on load (a 1.52pt table border is
 * 19304 EMU, which reads back as 2.026666666666667px), and the inputs that
 * showed them raw printed every digit of the float. PowerPoint shows two
 * decimals in its format panes, so the inspector does the same. Only the
 * displayed value is rounded: what the user types is committed as typed.
 */

/**
 * Round `value` to two decimals for an inspector input, or `fallback` (the
 * default the field shows) when the model holds no finite value.
 */
export function inspectorNumber(value: number | null | undefined, fallback = 0): number {
	const source = typeof value === 'number' && Number.isFinite(value) ? value : fallback;
	return Math.round(source * 100) / 100;
}
