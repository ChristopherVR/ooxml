/**
 * Standalone parser utilities for `a:rPr` (text run properties) XML nodes.
 *
 * The attributes and font slots are read by the shared DrawingML run codec
 * (`drawingml/text-run-codec.ts`, through `run-properties-neutral`), the
 * same rules the slide reader and the DOM reader run; these functions convert
 * the neutral values into {@link TextStyle} fields. The colour, outline and
 * hyperlink helpers below are simplified readers for the SmartArt projection.
 *
 * @module text-run-properties-parser
 */

import { drawingBooleanValue } from '../../../drawingml/text-paragraph-codec';
import type { TextStyle, XmlObject } from '../types';
import { extractColorChoiceXml } from './color-xml-preservation';
import {
	objectRunProperties,
	runAttributesFromDrawing,
	runFontMetadataFromDrawing,
} from './run-properties-neutral';

const EMU_PER_PX = 9525;

// ---------------------------------------------------------------------------
// Core run-property parsing
// ---------------------------------------------------------------------------

/**
 * Parse scalar run-level attributes from an `a:rPr` XML object.
 *
 * Handles: `@_sz`, `@_b`, `@_i`, `@_u`, `@_strike`, `@_cap`, `@_baseline`,
 * `@_kern`, `@_spc`, `@_lang`, `@_rtl`, `@_kumimoji`, `@_normalizeH`,
 * `@_noProof`, `@_dirty`, `@_err`, `@_smtClean`, `@_bmk`, `@_altLang`,
 * `@_smtId`. An authored `u="none"` or `cap="none"` is not recorded here (the
 * SmartArt projection has no writer that re-emits it).
 */
export function parseRunPropertyAttributes(rPr: XmlObject | undefined): TextStyle {
	if (!rPr) {
		return {};
	}
	const style: TextStyle = runAttributesFromDrawing(objectRunProperties(rPr), false);
	// The `@rtl` ATTRIBUTE is not part of CT_TextCharacterProperties (the run
	// flag is the `a:rtl` child); SDK-built content writes it, so it is read.
	const runRtl = drawingBooleanValue(rPr['@_rtl'] === undefined ? undefined : String(rPr['@_rtl']));
	if (runRtl !== undefined) {
		style.rtl = runRtl;
	}
	return style;
}

/**
 * Parse font family child elements from `a:rPr`.
 *
 * Extracts `a:latin`, `a:ea`, `a:cs` typeface attributes and their
 * CT_TextFont metadata. The Latin face is the primary family; per-script
 * faces are stored in their dedicated fields and selected at render time.
 */
export function parseRunFontElements(
	rPr: XmlObject | undefined,
): Pick<
	TextStyle,
	| 'fontFamily'
	| 'eastAsiaFont'
	| 'complexScriptFont'
	| 'latinFontPanose'
	| 'latinFontPitchFamily'
	| 'latinFontCharset'
	| 'eastAsiaFontPanose'
	| 'eastAsiaFontPitchFamily'
	| 'eastAsiaFontCharset'
	| 'complexScriptFontPanose'
	| 'complexScriptFontPitchFamily'
	| 'complexScriptFontCharset'
> {
	if (!rPr) {
		return {};
	}
	const { symbol: _symbol, ...run } = objectRunProperties(rPr);
	const {
		symbolFontPanose: _panose,
		symbolFontPitchFamily: _pitch,
		symbolFontCharset: _charset,
		...metadata
	} = runFontMetadataFromDrawing(run);
	return {
		...(run.latin ? { fontFamily: run.latin.typeface } : {}),
		...(run.eastAsian ? { eastAsiaFont: run.eastAsian.typeface } : {}),
		...(run.complexScript ? { complexScriptFont: run.complexScript.typeface } : {}),
		...metadata,
	};
}

/**
 * Parse underline colour from `a:uFill` or `a:uLn` child elements.
 *
 * Extracts a hex colour string from the solidFill inside `a:uFill` or
 * `a:uLn`.  This is a simplified extraction — the full runtime uses the
 * theme-aware `parseColor` method; this standalone function only handles
 * `a:srgbClr/@_val` for direct testing.
 */
export function parseRunUnderlineColor(rPr: XmlObject | undefined): string | undefined {
	if (!rPr) {
		return undefined;
	}

	const uFill = rPr['a:uFill'] as XmlObject | undefined;
	const uLn = rPr['a:uLn'] as XmlObject | undefined;
	const fillSource = uFill?.['a:solidFill'] || uLn?.['a:solidFill'];
	if (!fillSource) {
		return undefined;
	}

	const srgb = (fillSource as XmlObject)['a:srgbClr'] as XmlObject | undefined;
	if (srgb?.['@_val']) {
		return String(srgb['@_val']);
	}
	return undefined;
}

/**
 * Parse text outline from `a:ln` child element.
 */
export function parseRunTextOutline(
	rPr: XmlObject | undefined,
): Pick<TextStyle, 'textOutlineWidth' | 'textOutlineColor' | 'textOutlineDash'> {
	const result: Pick<TextStyle, 'textOutlineWidth' | 'textOutlineColor' | 'textOutlineDash'> = {};
	if (!rPr) {
		return result;
	}

	const textLn = rPr['a:ln'] as XmlObject | undefined;
	if (!textLn) {
		return result;
	}

	const textOutlineW = Number.parseInt(String(textLn['@_w'] || ''), 10);
	if (Number.isFinite(textOutlineW) && textOutlineW > 0) {
		result.textOutlineWidth = textOutlineW / EMU_PER_PX;
	}

	const dash = (textLn['a:prstDash'] as XmlObject | undefined)?.['@_val'];
	if (typeof dash === 'string' && dash !== 'solid') {
		result.textOutlineDash = dash;
	}

	const solidFill = textLn['a:solidFill'] as XmlObject | undefined;
	if (solidFill) {
		const srgb = solidFill['a:srgbClr'] as XmlObject | undefined;
		if (srgb?.['@_val']) {
			result.textOutlineColor = String(srgb['@_val']);
		}
	}

	return result;
}

/**
 * Parse hyperlink from `a:hlinkClick` child element.
 *
 * Extracts relationship ID, tooltip, action, and additional attributes.
 * Does NOT resolve the relationship target (requires a relationship map).
 */
export function parseRunHyperlink(
	rPr: XmlObject | undefined,
): Pick<
	TextStyle,
	| 'hyperlinkRId'
	| 'hyperlinkTooltip'
	| 'hyperlinkAction'
	| 'hyperlinkInvalidUrl'
	| 'hyperlinkTargetFrame'
	| 'hyperlinkHistory'
	| 'hyperlinkHighlightClick'
	| 'hyperlinkEndSound'
> {
	const result: Pick<
		TextStyle,
		| 'hyperlinkRId'
		| 'hyperlinkTooltip'
		| 'hyperlinkAction'
		| 'hyperlinkInvalidUrl'
		| 'hyperlinkTargetFrame'
		| 'hyperlinkHistory'
		| 'hyperlinkHighlightClick'
		| 'hyperlinkEndSound'
	> = {};
	if (!rPr) {
		return result;
	}

	const hlinkClick = rPr['a:hlinkClick'] as XmlObject | undefined;
	if (!hlinkClick) {
		return result;
	}

	const rId = String(hlinkClick['@_r:id'] || hlinkClick['@_id'] || '').trim();
	if (rId.length > 0) {
		result.hyperlinkRId = rId;
	}

	const tooltip = String(hlinkClick['@_tooltip'] || '').trim();
	if (tooltip) {
		result.hyperlinkTooltip = tooltip;
	}

	const action = String(hlinkClick['@_action'] || '').trim();
	if (action) {
		result.hyperlinkAction = action;
	}

	const invalidUrl = String(hlinkClick['@_invalidUrl'] || '').trim();
	if (invalidUrl) {
		result.hyperlinkInvalidUrl = invalidUrl;
	}

	const tgtFrame = String(hlinkClick['@_tgtFrame'] || '').trim();
	if (tgtFrame) {
		result.hyperlinkTargetFrame = tgtFrame;
	}

	if (hlinkClick['@_history'] !== undefined) {
		const hVal = String(hlinkClick['@_history']).trim().toLowerCase();
		result.hyperlinkHistory = hVal !== '0' && hVal !== 'false';
	}

	if (hlinkClick['@_highlightClick'] !== undefined) {
		const hcVal = String(hlinkClick['@_highlightClick']).trim().toLowerCase();
		result.hyperlinkHighlightClick = hcVal === '1' || hcVal === 'true';
	}

	if (hlinkClick['@_endSnd'] !== undefined) {
		const esVal = String(hlinkClick['@_endSnd']).trim().toLowerCase();
		result.hyperlinkEndSound = esVal === '1' || esVal === 'true';
	}

	return result;
}

/**
 * Parse solidFill colour from `a:solidFill` child element.
 * Only handles direct `a:srgbClr` for standalone use.
 */
export function parseRunSolidFillColor(rPr: XmlObject | undefined): string | undefined {
	if (!rPr) {
		return undefined;
	}
	const solidFill = rPr['a:solidFill'] as XmlObject | undefined;
	if (!solidFill) {
		return undefined;
	}

	const srgb = solidFill['a:srgbClr'] as XmlObject | undefined;
	if (srgb?.['@_val']) {
		return String(srgb['@_val']);
	}
	return undefined;
}

/**
 * Extract the raw colour-choice XML node from `a:rPr/a:solidFill` for
 * round-trip preservation. Returns the wrapping single-key object (e.g.
 * `{ 'a:schemeClr': {...} }`) so the save layer can re-emit it verbatim
 * when the user did not edit the colour.
 */
export function parseRunSolidFillColorXml(rPr: XmlObject | undefined): XmlObject | undefined {
	if (!rPr) {
		return undefined;
	}
	const solidFill = rPr['a:solidFill'] as XmlObject | undefined;
	if (!solidFill) {
		return undefined;
	}
	return extractColorChoiceXml(solidFill);
}

/**
 * Parse symbol font from `a:sym` child element.
 */
export function parseRunSymbolFont(rPr: XmlObject | undefined): string | undefined {
	return rPr ? objectRunProperties(rPr).symbol?.typeface : undefined;
}

/**
 * Parse `a:sym` font metadata (panose / pitchFamily / charset).
 *
 * Returned values are written back on save to keep the round-trip stable
 * for fonts where these attributes were authored.
 */
export function parseRunSymbolFontMetadata(
	rPr: XmlObject | undefined,
): Pick<TextStyle, 'symbolFontPanose' | 'symbolFontPitchFamily' | 'symbolFontCharset'> {
	if (!rPr) {
		return {};
	}
	const symbol = objectRunProperties(rPr).symbol;
	const { symbolFontPanose, symbolFontPitchFamily, symbolFontCharset } = runFontMetadataFromDrawing(
		symbol ? { symbol } : {},
	);
	return {
		...(symbolFontPanose !== undefined ? { symbolFontPanose } : {}),
		...(symbolFontPitchFamily !== undefined ? { symbolFontPitchFamily } : {}),
		...(symbolFontCharset !== undefined ? { symbolFontCharset } : {}),
	};
}
