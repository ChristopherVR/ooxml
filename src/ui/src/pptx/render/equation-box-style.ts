/**
 * Box style of a PURE equation shape (OMML and nothing else), shared by every
 * binding.
 *
 * The equation box honours the shape's own text body like any text box: the
 * insets, the anchor (`bodyPr/@anchor`, top by default) and the autofit come
 * from `getTextBlockStyle`, which makes the box a column flex container whose
 * `justify-content` is the anchor. Only the cross axis is overridden so each
 * equation row is centred horizontally. A binding that built its own wrapper
 * (a row flex with `align-items: center`) centred the maths vertically in the
 * box instead of anchoring it, which put the equation about 20px lower than
 * the vanilla, React and Angular renderers on Mathematical_Equations slide 5.
 */

/** Font stack of rendered maths, the same one the inline equation run uses. */
export const EQUATION_FONT_FAMILY = "'Cambria Math', 'STIX Two Math', serif";

type StyleValue = string | number | undefined;

/**
 * Wrapper style for a pure equation shape: the shape's text-block style (as a
 * property map) with the cross axis centred, no wrapping, and the maths font.
 */
export function buildEquationBlockStyle<T extends object>(
	textBlockStyle: T,
): T & Record<string, StyleValue> {
	return {
		...textBlockStyle,
		flexWrap: 'nowrap',
		alignItems: 'center',
		gap: '0.25em',
		fontFamily: EQUATION_FONT_FAMILY,
	};
}
