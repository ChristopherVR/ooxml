/**
 * table-style-effect-ref.ts - resolve a table style's `a:tblBg/a:effectRef`
 * ("Themed Style 1/2") against the theme's effect styles.
 *
 * PowerPoint draws the referenced effect (an outer shadow in every Office
 * theme's third effect style) under the whole table, on the table's outline.
 *
 * @module table-style-effect-ref
 */
import { applyDrawingColorTransforms } from '../../color';
import type {
	ParsedTableBackgroundShadow,
	ParsedTableFillRef,
	PptxThemeFormatScheme,
	XmlObject,
} from '../../types';

/** `#RRGGBB` from a stored colour, or `undefined` when it is not a plain hex colour. */
function hexOf(color: string | undefined): string | undefined {
	return color && /^#[0-9a-f]{6}$/iu.test(color) ? color : undefined;
}

/**
 * The outer shadow `effectRef` selects from `formatScheme.effectStyles`
 * (`idx` 1-3), or `undefined` when there is none (no theme, no style at that
 * index, an empty effect list). A shadow whose colour is the `phClr`
 * placeholder takes the reference's own colour (`colorMap` maps scheme keys to
 * hex), with the shadow colour's transforms (usually just alpha) applied.
 */
export function resolveTableBackgroundEffectRef(
	effectRef: ParsedTableFillRef,
	formatScheme: PptxThemeFormatScheme | undefined,
	colorMap: Record<string, string> | undefined,
): ParsedTableBackgroundShadow | undefined {
	const effect = effectRef.idx >= 1 ? formatScheme?.effectStyles[effectRef.idx - 1] : undefined;
	if (!effect) {
		return undefined;
	}
	let color = hexOf(effect.shadowColor);
	if (!color) {
		const overrideKey = effectRef.color?.schemeColor;
		const base = overrideKey && colorMap ? colorMap[overrideKey] : effectRef.color?.color;
		const lst = (effect.rawNode as XmlObject | undefined)?.['a:effectLst'] as XmlObject | undefined;
		const shadow = lst?.['a:outerShdw'] as XmlObject | undefined;
		const colorNode = shadow?.['a:schemeClr'] as XmlObject | undefined;
		if (!shadow || !base) {
			return undefined;
		}
		color = hexOf(colorNode ? applyDrawingColorTransforms(base, colorNode) : base) ?? hexOf(base);
	}
	if (!color) {
		return undefined;
	}
	return {
		color,
		opacity: effect.shadowOpacity ?? 1,
		blur: effect.shadowBlur ?? 0,
		offsetX: effect.shadowOffsetX ?? 0,
		offsetY: effect.shadowOffsetY ?? 0,
	};
}
