/**
 * The pptx object-tree adapter of the shared DrawingML run property codec
 * (`drawingml/text-run-codec.ts`, docs/agnostic-core-plan.md step 3).
 *
 * The rules for the `a:rPr` / `a:defRPr` / `a:endParaRPr` attributes, the
 * font slots, the run `a:rtl` and the underline follow-text markers are the
 * codec's, the same ones the DOM reader and writer run. This module reads an
 * object-tree node through the codec and converts the neutral values into the
 * pptx {@link TextStyle} fields (points to px, fractions to the file's
 * thousandths of a percent, the PowerPoint underline, strike and caps flags),
 * and back for the writer. Colours, fills, outlines, effects and hyperlinks
 * stay with their own codecs.
 *
 * @module run-properties-neutral
 */
import type { DrawingFont, DrawingRunProperties } from '../../../drawingml/text-model';
import { readDrawingRunCodecProperties } from '../../../drawingml/text-run-codec';
import type { TextStyle } from '../types';
import type { UnderlineStyle } from '../types/common';
import { objectTextElementView } from './paragraph-properties-neutral';

const PX_PER_POINT = 96 / 72;

/** The codec families of an `a:rPr` / `a:defRPr` / `a:endParaRPr` object node; empty when absent. */
export function objectRunProperties(rPr: unknown): DrawingRunProperties {
	const view = objectTextElementView(rPr);
	return view ? readDrawingRunCodecProperties(view) : {};
}

/** The run attribute fields of {@link TextStyle} that every pptx run reader shares. */
export type RunAttributeFields = Pick<
	TextStyle,
	| 'fontSize'
	| 'bold'
	| 'italic'
	| 'underline'
	| 'underlineStyle'
	| 'underlineExplicitNone'
	| 'strikethrough'
	| 'strikeType'
	| 'textCaps'
	| 'textCapsExplicitNone'
	| 'baseline'
	| 'characterSpacing'
	| 'kerning'
	| 'language'
	| 'altLanguage'
	| 'kumimoji'
	| 'normalizeHeight'
	| 'noProof'
	| 'dirty'
	| 'spellingError'
	| 'smartTagClean'
	| 'smartTagId'
	| 'bookmark'
>;

const OFF_TOKENS = new Set(['none', '0', 'false']);

/**
 * The run attribute fields of neutral run properties. `explicitNone` records
 * an authored `u="none"` / `cap="none"` (the slide reader keeps them so the
 * writer re-emits the suppression; the SmartArt projection does not).
 */
export function runAttributesFromDrawing(
	run: DrawingRunProperties,
	explicitNone = true,
): RunAttributeFields {
	const out: RunAttributeFields = {};
	if (run.sizePt !== undefined) out.fontSize = run.sizePt * PX_PER_POINT;
	if (run.bold !== undefined) out.bold = run.bold;
	if (run.italic !== undefined) out.italic = run.italic;
	if (run.underline !== undefined) {
		const token = run.underline.toLowerCase();
		out.underline = !OFF_TOKENS.has(token);
		if (out.underline) {
			out.underlineStyle = run.underline as UnderlineStyle;
		} else if (explicitNone && token === 'none') {
			// Explicit `<a:rPr u="none"/>` suppression, re-emitted by the writer
			// instead of collapsing to inherit.
			out.underlineExplicitNone = true;
		}
	}
	if (run.strike !== undefined) {
		const token = run.strike.toLowerCase();
		out.strikethrough = token !== 'nostrike' && !OFF_TOKENS.has(token);
		if (out.strikethrough) {
			out.strikeType = token === 'dblstrike' ? 'dblStrike' : 'sngStrike';
		}
	}
	const caps = run.caps?.toLowerCase();
	if (caps === 'all' || caps === 'small') {
		out.textCaps = caps;
	} else if (explicitNone && caps === 'none') {
		out.textCaps = 'none';
		out.textCapsExplicitNone = true;
	}
	// An authored `baseline="0"` is kept as 0 (not collapsed to unset): two
	// runs differing only by it must not compare equal, or the save path merges
	// them into one run.
	if (run.baseline !== undefined) out.baseline = Math.round(run.baseline * 100000);
	if (run.spacingPt !== undefined) out.characterSpacing = Math.round(run.spacingPt * 100);
	if (run.kernPt !== undefined) out.kerning = Math.round(run.kernPt * 100);
	if (run.language) out.language = run.language;
	if (run.altLanguage) out.altLanguage = run.altLanguage;
	if (run.kumimoji !== undefined) out.kumimoji = run.kumimoji;
	if (run.normalizeHeight !== undefined) out.normalizeHeight = run.normalizeHeight;
	if (run.noProof !== undefined) out.noProof = run.noProof;
	if (run.dirty !== undefined) out.dirty = run.dirty;
	if (run.spellingError !== undefined) out.spellingError = run.spellingError;
	if (run.smartTagClean !== undefined) out.smartTagClean = run.smartTagClean;
	if (run.smartTagId !== undefined) out.smartTagId = run.smartTagId;
	if (run.bookmark) out.bookmark = run.bookmark;
	return out;
}

type FontPrefix = 'latinFont' | 'eastAsiaFont' | 'complexScriptFont' | 'symbolFont';

const FONT_PREFIXES = [
	['latin', 'latinFont'],
	['eastAsian', 'eastAsiaFont'],
	['complexScript', 'complexScriptFont'],
	['symbol', 'symbolFont'],
] as const satisfies ReadonlyArray<readonly [keyof DrawingRunProperties, FontPrefix]>;

/** The `@panose` / `@pitchFamily` / `@charset` fields of the four font slots. */
export function runFontMetadataFromDrawing(run: DrawingRunProperties): Partial<TextStyle> {
	const out: Partial<TextStyle> = {};
	for (const [key, prefix] of FONT_PREFIXES) {
		const font = run[key] as DrawingFont | undefined;
		if (!font) continue;
		if (font.panose) (out as Record<string, unknown>)[`${prefix}Panose`] = font.panose;
		if (font.pitchFamily !== undefined)
			(out as Record<string, unknown>)[`${prefix}PitchFamily`] = font.pitchFamily;
		if (font.charset !== undefined)
			(out as Record<string, unknown>)[`${prefix}Charset`] = font.charset;
	}
	return out;
}

/** A font slot of a {@link TextStyle} on the neutral model (typeface plus its hints). */
export function textFontToDrawing(
	typeface: string,
	panose: string | undefined,
	pitchFamily: number | undefined,
	charset: number | undefined,
): DrawingFont {
	return {
		typeface,
		...(panose ? { panose } : {}),
		...(typeof pitchFamily === 'number' && Number.isFinite(pitchFamily) ? { pitchFamily } : {}),
		...(typeof charset === 'number' && Number.isFinite(charset) ? { charset } : {}),
	};
}
