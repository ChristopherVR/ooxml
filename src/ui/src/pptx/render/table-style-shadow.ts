/**
 * The drop shadow a table style draws under the whole table.
 *
 * "Themed Style 1" and "Themed Style 2" reference a theme effect style from
 * `a:tblBg/a:effectRef` (the third one is an outer shadow in every Office
 * theme); PowerPoint paints it on the table's outline, which no cell can show,
 * so it belongs on the `<table>` element (see `tableContainerCss`). The
 * shadow takes the same blur profile as a shape's outer shadow
 * (`effect-profile.ts`).
 *
 * @module render/table-style-shadow
 */
import type { ParsedTableStyleEntry } from 'ooxml-core/pptx';

import { getOuterShadowCss } from './visual-effects';

/**
 * The CSS `box-shadow` of the table style's background effect, or `undefined`
 * when the style has none (or `entry` has not had its theme references
 * resolved, see `resolveTableStyleThemeRefs`).
 */
export function tableStyleShadowCss(entry: ParsedTableStyleEntry | undefined): string | undefined {
	const shadow = entry?.tableBackground?.shadow;
	if (!shadow) {
		return undefined;
	}
	return getOuterShadowCss({
		shadowColor: shadow.color,
		shadowOpacity: shadow.opacity,
		shadowBlur: shadow.blur,
		shadowOffsetX: shadow.offsetX,
		shadowOffsetY: shadow.offsetY,
		// Marks the values as DrawingML's (blur radius, not a CSS blur length).
		outerShadowXml: {},
	});
}
