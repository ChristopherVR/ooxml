import { resolveDrawingColor } from '../../../../drawingml/drawing-color';
import { parseThemeFormatScheme } from '../../../../drawingml/format-scheme';
import type { ThemeFormatScheme } from '../../../../drawingml/format-scheme';
import type { DrawingColor } from '../../../../drawingml/types';
import { NS, type XmlElement } from '../../../../xml/index';
import { XmlObject } from '../../types';
import { withThemePlaceholderColor } from '../../utils/theme-style-reference';
import type {
	PptxThemeEffectStyle,
	PptxThemeFillStyle,
	PptxThemeFormatScheme,
	PptxThemeLineStyle,
} from '../../types';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeLayoutPreviews';
import {
	objectNodeElement,
	themeFormatSchemeView,
	type ThemeStyleViewContext,
} from './theme-format-scheme-view';

const EMPTY_SCHEME: ThemeFormatScheme = {
	fillStyles: [],
	lineStyles: [],
	effectStyles: [],
	backgroundFillStyles: [],
};

/**
 * Reading the theme format scheme (`a:fmtScheme`): the shared DrawingML reader
 * (`parseThemeFormatScheme`) reads the element and {@link themeFormatSchemeView} projects it onto
 * the pptx view, resolving colours through the runtime's theme.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * The projection context, colours resolved through the active theme and colour map. Neutral
	 * colours arrive with `phClr` already substituted (`resolveTheme*Style`); `placeholder`
	 * (`#RRGGBB`) stands for it in what the object tree still holds (gradients, 3D colours).
	 */
	protected themeStyleViewContext(placeholder?: string): ThemeStyleViewContext {
		// The gradient codec reads the object tree and resolves `phClr` through the theme map.
		const withPlaceholder = <T>(read: () => T): T =>
			placeholder ? withThemePlaceholderColor(this.themeColorMap, placeholder, read) : read();
		return {
			color: (color) => this.resolveThemeDrawingColor(color),
			opacity: (holder) => this.extractColorOpacity(holder),
			gradient: (node) =>
				withPlaceholder(() => ({
					color: this.extractGradientFillColor(node),
					opacity: this.extractGradientOpacity(node),
					gradientStops: this.extractGradientStops(node),
					gradientAngle: this.extractGradientAngle(node),
					gradientType: this.extractGradientType(node),
					gradientCss: this.extractGradientFillCss(node),
				})),
			gradientColor: (node) => withPlaceholder(() => this.extractGradientFillColor(node)),
			objectColor: (holder) => this.parseColor(holder, placeholder),
			normalizeDash: (value) => this.normalizeStrokeDashType(value),
			emuPerPx: PptxHandlerRuntime.EMU_PER_PX,
		};
	}

	/**
	 * `#RRGGBB` of a neutral colour through the active theme: a scheme name resolves through the
	 * runtime (colour map and overrides), as the object-tree codec resolves it.
	 */
	protected resolveThemeDrawingColor(color: DrawingColor | undefined): string | undefined {
		if (!color) {
			return undefined;
		}
		return resolveDrawingColor(color, {
			scheme: (name) => {
				const key = name.toLowerCase();
				return key ? this.resolveThemeColor(key) : undefined;
			},
		})?.hex;
	}

	/**
	 * Read an `a:fmtScheme` into the pptx view. `source` is the same element in the DOM, when the
	 * caller has the part's text (it keeps the document order of mixed fill lists); otherwise the
	 * object tree is written out and read back, which groups same-name siblings.
	 */
	protected readFormatScheme(fmtScheme: XmlObject, source?: XmlElement): PptxThemeFormatScheme {
		const element =
			source && source.namespaceURI === NS.a && source.localName === 'fmtScheme'
				? source
				: objectNodeElement((tree) => this.builder.build(tree), 'a:fmtScheme', fmtScheme);
		const scheme = parseThemeFormatScheme(element) ?? EMPTY_SCHEME;
		return themeFormatSchemeView(scheme, fmtScheme, this.themeStyleViewContext());
	}

	/** The fill list (`a:fillStyleLst` or `a:bgFillStyleLst`) of a format scheme, read alone. */
	protected parseFillStyleList(listNode: XmlObject | undefined): PptxThemeFillStyle[] {
		return listNode ? this.readFormatScheme({ 'a:fillStyleLst': listNode }).fillStyles : [];
	}

	/** The `a:lnStyleLst` of a format scheme, read alone. */
	protected parseLineStyleList(listNode: XmlObject | undefined): PptxThemeLineStyle[] {
		return listNode ? this.readFormatScheme({ 'a:lnStyleLst': listNode }).lineStyles : [];
	}

	/** The `a:effectStyleLst` of a format scheme, read alone. */
	protected parseEffectStyleList(listNode: XmlObject | undefined): PptxThemeEffectStyle[] {
		return listNode ? this.readFormatScheme({ 'a:effectStyleLst': listNode }).effectStyles : [];
	}
}
