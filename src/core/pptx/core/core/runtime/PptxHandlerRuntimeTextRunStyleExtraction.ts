import { themeColorRefFromColorChoice } from '../../color/theme-color-ref';
import { TextStyle, XmlObject } from '../../types';
import { extractColorChoiceXml } from '../../utils/color-xml-preservation';
import { resolveTextFillBlip } from '../../utils/text-fill-blip';
import {
	objectRunProperties,
	runAttributesFromDrawing,
	runFontMetadataFromDrawing,
} from '../../utils/run-properties-neutral';
import { xmlChild } from '../../utils/xml-access';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeTextRunEffects';

/**
 * The `EG_FillProperties` members `CT_TextCharacterProperties` may carry, minus
 * `a:noFill`. Any one of them is a run declaring a fill of its OWN, which
 * overrides an inherited `<a:noFill/>` from a lower style layer.
 */
const TEXT_FILL_ELEMENTS = [
	'a:solidFill',
	'a:gradFill',
	'a:pattFill',
	'a:blipFill',
	'a:grpFill',
] as const;

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	protected extractTextRunStyle(
		runProperties: XmlObject | undefined,
		align: TextStyle['align'],
		relationshipMap?: Map<string, string>,
		includeDefaultAlignment: boolean = true,
		slidePath?: string,
	): TextStyle {
		const style: TextStyle = includeDefaultAlignment ? { align } : {};
		if (!runProperties) {
			return style;
		}

		// The attributes (size, the `xsd:boolean` flags in every lexical form,
		// underline, strike, caps, baseline, spacing, kerning, language and the
		// metadata), the font slots, the run `a:rtl` and the underline
		// follow-text markers are read by the shared DrawingML run codec
		// (`run-properties-neutral`). Colours, fills, the outline, effects and
		// hyperlinks are read below by their own codecs.
		const neutral = objectRunProperties(runProperties);
		Object.assign(style, runAttributesFromDrawing(neutral));
		// Underline colour (a:uFill > a:solidFill or a:uLn > a:solidFill)
		const uFill = runProperties['a:uFill'] as XmlObject | undefined;
		const uLn = runProperties['a:uLn'] as XmlObject | undefined;
		const underlineColorSource = uFill?.['a:solidFill'] || uLn?.['a:solidFill'];
		if (underlineColorSource) {
			const underlineColor = this.parseColor(underlineColorSource as XmlObject);
			if (underlineColor) {
				style.underlineColor = underlineColor;
			}
		}
		// #85: Underline line styling (a:uLn width/dash/cap/ends) and the
		// follow-text markers (a:uLnTx / a:uFillTx). Previously only the uLn
		// solidFill colour above was read; the line properties were dropped.
		if (uLn) {
			const line: NonNullable<TextStyle['underlineLine']> = {};
			const widthEmu = Number.parseInt(String(uLn['@_w'] ?? ''), 10);
			if (Number.isFinite(widthEmu)) {
				line.widthEmu = widthEmu;
			}
			const compound = String(uLn['@_cmpd'] ?? '').trim();
			if (compound) {
				line.compound = compound;
			}
			const cap = String(uLn['@_cap'] ?? '').trim();
			if (cap) {
				line.cap = cap;
			}
			const algn = String(uLn['@_algn'] ?? '').trim();
			if (algn) {
				line.algn = algn;
			}
			const prstDash = String((uLn['a:prstDash'] as XmlObject | undefined)?.['@_val'] ?? '').trim();
			if (prstDash) {
				line.prstDash = prstDash;
			}
			const headEnd = uLn['a:headEnd'];
			if (headEnd && typeof headEnd === 'object') {
				line.headEndXml = headEnd as XmlObject;
			}
			const tailEnd = uLn['a:tailEnd'];
			if (tailEnd && typeof tailEnd === 'object') {
				line.tailEndXml = tailEnd as XmlObject;
			}
			if (Object.keys(line).length > 0) {
				style.underlineLine = line;
			}
		}
		if (neutral.underlineLineFollowsText) {
			style.underlineLineFollowsText = true;
		}
		if (neutral.underlineFillFollowsText) {
			style.underlineFillFollowsText = true;
		}
		// Text outline (a:rPr > a:ln)
		const textLn = runProperties['a:ln'] as XmlObject | undefined;
		if (textLn) {
			const textOutlineW = Number.parseInt(String(textLn['@_w'] || ''), 10);
			if (Number.isFinite(textOutlineW) && textOutlineW > 0) {
				style.textOutlineWidth = textOutlineW / PptxHandlerRuntime.EMU_PER_PX;
			}
			const textOutlineDash = (textLn['a:prstDash'] as XmlObject | undefined)?.['@_val'];
			if (typeof textOutlineDash === 'string' && textOutlineDash !== 'solid') {
				style.textOutlineDash = textOutlineDash;
			}
			const textOutlineFill = textLn['a:solidFill'] as XmlObject | undefined;
			if (textOutlineFill) {
				const outlineColor = this.parseColor(textOutlineFill);
				if (outlineColor) {
					style.textOutlineColor = outlineColor;
				}
			}
		}
		// No fill on text run (a:rPr > a:noFill): hollow / outline-only text.
		//
		// This flag has to be able to say `false` as well as `true`. Run styles are
		// assembled as `{...mergedDefaultRunStyle, ...extractTextRunStyle(rPr)}`,
		// where the lower layers are themselves `extractTextRunStyle` of the
		// `a:lstStyle` levels, the layout / master `a:defRPr` and `a:endParaRPr`.
		// Only ever setting it `true` meant a run that OVERRODE an inherited
		// `<a:noFill/>` with a fill of its own left the slot untouched, so the
		// inherited `true` survived the spread: the run rendered hollow
		// (`hollowTextFillStyle` in shared) and the writer re-emitted `<a:noFill/>`
		// over its real fill, losing the colour permanently on the first save.
		// Recording the override explicitly is what lets the spread clear it.
		if (runProperties['a:noFill'] !== undefined) {
			style.textFillNone = true;
		} else if (TEXT_FILL_ELEMENTS.some((name) => runProperties[name] !== undefined)) {
			style.textFillNone = false;
		}
		// Text highlight colour
		if (runProperties['a:highlight']) {
			const highlightNode = xmlChild(runProperties, 'a:highlight');
			const highlightHex = this.parseColor(highlightNode);
			if (highlightHex) {
				style.highlightColor = highlightHex;
			}
			// Preserve the original colour-choice so a themed (`a:schemeClr`)
			// highlight re-emits as itself rather than flattening to srgbClr.
			const highlightXml = extractColorChoiceXml(highlightNode);
			if (highlightXml) {
				style.highlightColorXml = highlightXml;
			}
		}
		// Text fill variants (gradient/pattern on a:rPr)
		const textFillVariants = this.extractTextFillVariants(runProperties);
		if (textFillVariants.textFillGradient) {
			style.textFillGradient = textFillVariants.textFillGradient;
			style.textFillGradientStops = textFillVariants.textFillGradientStops;
			style.textFillGradientAngle = textFillVariants.textFillGradientAngle;
			style.textFillGradientType = textFillVariants.textFillGradientType;
		}
		if (textFillVariants.textFillPattern) {
			style.textFillPattern = textFillVariants.textFillPattern;
			style.textFillPatternForeground = textFillVariants.textFillPatternForeground;
			style.textFillPatternBackground = textFillVariants.textFillPatternBackground;
		}
		if (textFillVariants.textFillBlipXml) {
			style.textFillBlipXml = textFillVariants.textFillBlipXml;
			const rels = slidePath ? this.slideRelsMap.get(slidePath) : undefined;
			const resolved =
				slidePath && rels
					? resolveTextFillBlip(
							textFillVariants.textFillBlipXml,
							(relId) => rels.get(relId),
							(target) => this.resolveImagePath(slidePath, target),
							this.allowExternalImages === true,
						)
					: undefined;
			if (resolved) {
				style.textFillBlipUrl = resolved.url;
				style.textFillBlipMode = resolved.mode;
			}
		}
		// Run-level right-to-left. On CT_TextCharacterProperties `rtl` is a child
		// ELEMENT of type CT_Boolean (`<a:rtl val="1"/>`) whose `@val` defaults
		// to true when omitted; the ATTRIBUTE spelling belongs to
		// CT_TextParagraphProperties. Reading only `@_rtl` here meant an
		// authored `<a:rtl/>` never loaded at all. The attribute is still read
		// as a fallback for SDK-built content that put it there.
		const runRtl = neutral.rightToLeft ?? this.parseOptionalBooleanAttr(runProperties['@_rtl']);
		if (runRtl !== undefined) {
			style.rtl = runRtl;
		}

		const latinTypefaceToken = neutral.latin?.typeface;
		const eaTypefaceToken = neutral.eastAsian?.typeface;
		const csTypefaceToken = neutral.complexScript?.typeface;
		// The run's primary CSS family is its Latin face. East Asian and complex
		// script faces are preserved below and applied only to matching glyphs.
		const chosenTypeface = latinTypefaceToken;
		const resolvedTypeface = this.resolveThemeTypeface(chosenTypeface);
		if (resolvedTypeface) {
			style.fontFamily = resolvedTypeface;
		}

		// #84: Preserve the theme-font token (`+mj-lt` / `+mn-ea` / ...) on the
		// typeface it was authored on, so the writer can re-emit the token rather
		// than the flattened concrete face. `fontFamily` above still holds the
		// resolved concrete face for rendering.
		if (latinTypefaceToken && latinTypefaceToken.startsWith('+')) {
			style.latinFontThemeToken = latinTypefaceToken;
		}
		if (eaTypefaceToken && eaTypefaceToken.startsWith('+')) {
			style.eastAsiaFontThemeToken = eaTypefaceToken;
		}
		if (csTypefaceToken && csTypefaceToken.startsWith('+')) {
			style.complexScriptFontThemeToken = csTypefaceToken;
		}

		// Store per-script font families for Unicode font fallback. Only set when
		// the source actually authored an `a:ea` / `a:cs` typeface; otherwise the
		// writer must not synthesize one (see #84).
		const eaTypeface = this.resolveThemeTypeface(eaTypefaceToken);
		if (eaTypeface) {
			style.eastAsiaFont = eaTypeface;
		}
		const csTypeface = this.resolveThemeTypeface(csTypefaceToken);
		if (csTypeface) {
			style.complexScriptFont = csTypeface;
		}

		const solidFill = xmlChild(runProperties, 'a:solidFill');
		if (solidFill) {
			style.color = this.parseColor(solidFill);
			const colorXml = extractColorChoiceXml(solidFill);
			if (colorXml) {
				style.colorXml = colorXml;
			}
			const colorRef = themeColorRefFromColorChoice(solidFill);
			if (colorRef) {
				style.colorRef = colorRef;
			}
		}

		// Hyperlinks (a:hlinkClick, a:hlinkMouseOver)
		this.applyHyperlinkStyle(style, runProperties, relationshipMap);

		// Symbol font (a:sym)
		if (neutral.symbol) {
			style.symbolFont = neutral.symbol.typeface;
		}

		// Per-script font metadata (CT_TextFont @panose, @pitchFamily, @charset).
		Object.assign(style, runFontMetadataFromDrawing(neutral));

		// Text run effects (a:effectLst on a:rPr)
		const runEffectList = runProperties['a:effectLst'] as XmlObject | undefined;
		if (runEffectList && typeof runEffectList === 'object') {
			this.applyTextRunEffects(style, runEffectList);
		}
		if (
			includeDefaultAlignment &&
			runEffectList !== undefined &&
			(typeof runEffectList !== 'object' || Object.keys(runEffectList).length === 0)
		) {
			style.textEffectsExplicitNone = true;
		}

		// Text run effect graph (a:effectDag on a:rPr): ECMA-376
		// §21.1.2.3.6 allows `effectDag` as an alternative to `effectLst`.
		this.applyTextRunEffectDag(style, runProperties);

		// Run-level `a:extLst` preserved verbatim so authored extensions survive
		// a round-trip. Captured only for real runs (`includeDefaultAlignment` is
		// true only on run/field/direct-text extraction, false for the
		// defRPr / level / endParaRPr default-style passes) so a default style's
		// extLst does not propagate onto every run of the paragraph.
		if (includeDefaultAlignment) {
			const runExtLst = runProperties['a:extLst'];
			if (runExtLst && typeof runExtLst === 'object') {
				style.runPropertiesExtLstXml = runExtLst as XmlObject;
			}
		}

		return style;
	}
}
