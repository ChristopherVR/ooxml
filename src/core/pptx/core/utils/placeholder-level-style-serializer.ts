/**
 * Serialize a {@link PlaceholderTextLevelStyle} back into an `a:lvlXpPr` /
 * `a:defPPr` node (CT_TextParagraphProperties, ECMA-376 §21.1.2.2.7), the
 * exact inverse of `parsePlaceholderLevelStyle`
 * (`PptxHandlerRuntimePlaceholderStyles.ts`).
 *
 * Every write merges into the existing node via {@link mergeOrderedXml}
 * rather than rebuilding it: only the fields explicitly set on the typed
 * style are touched, and any attribute, run-property, or child this model
 * does not cover (extensions, `a:buBlip`, ...) survives untouched, in its
 * original schema position. Preserved colour nodes (`colorChoiceXml`,
 * `bulletColorXml`) are re-emitted verbatim so theme aliases survive.
 *
 * @module placeholder-level-style-serializer
 */
import type { DrawingBullet, DrawingParagraphProperties } from '../../../drawingml/text-model';
import {
	drawingBulletChild,
	drawingBulletSizeChild,
	drawingParagraphAttributeEntries,
} from '../../../drawingml/text-paragraph-codec';
import type { PlaceholderTextLevelStyle, XmlObject } from '../types';
import { colorsEqual } from './color-xml-preservation';
import { mergeOrderedXml } from './ordered-xml-merge';
import { objectAttributes, objectTextSpacing } from './paragraph-properties-neutral';
import { serializeTabStop } from './tab-stops';

const EMU_PER_PX = 9525;
/** px -> hundredths-of-a-point: px * (72/96) * 100. */
const PT_HUNDREDTHS_PER_PX = 75;

/** `TextStyle['align']` token -> `ST_TextAlignType`; the last three are spelled the same on both sides. */
const ALIGN_TO_XML: Record<string, string> = {
	left: 'l',
	center: 'ctr',
	right: 'r',
	justify: 'just',
	justLow: 'justLow',
	dist: 'dist',
	thaiDist: 'thaiDist',
};

/** CT_TextParagraphProperties child order (bullet group: buClr*, buSz*, buFont*, bu<type>). */
const LEVEL_PPR_CHILD_ORDER = [
	'a:lnSpc',
	'a:spcBef',
	'a:spcAft',
	'a:buClr',
	'a:buSzPct',
	'a:buSzPts',
	'a:buFont',
	'a:buNone',
	'a:buAutoNum',
	'a:buChar',
	'a:tabLst',
	'a:defRPr',
	'a:extLst',
] as const;

/** CT_TextCharacterProperties child order (only the fill-choice group needs mutual exclusion here). */
export const DEF_RPR_CHILD_ORDER = [
	'a:ln',
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
	'a:effectLst',
	'a:effectDag',
	'a:highlight',
	'a:uLnTx',
	'a:uLn',
	'a:uFillTx',
	'a:uFill',
	'a:latin',
	'a:ea',
	'a:cs',
	'a:sym',
	'a:hlinkClick',
	'a:hlinkMouseOver',
	'a:rtl',
	'a:extLst',
] as const;

const FILL_VARIANTS = ['a:noFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'] as const;

/** The paragraph attributes of a level style on the neutral model. */
function levelAttributesToDrawing(style: PlaceholderTextLevelStyle): DrawingParagraphProperties {
	const out: DrawingParagraphProperties = {};
	if (style.alignment !== undefined) {
		out.align = ALIGN_TO_XML[style.alignment] ?? style.alignment;
	}
	if (style.marginLeft !== undefined) out.marginLeftEmu = style.marginLeft * EMU_PER_PX;
	if (style.marginRight !== undefined) out.marginRightEmu = style.marginRight * EMU_PER_PX;
	if (style.indent !== undefined) out.indentEmu = style.indent * EMU_PER_PX;
	if (style.rtl !== undefined) out.rightToLeft = style.rtl;
	if (style.defaultTabSize !== undefined) {
		out.defaultTabSizeEmu = style.defaultTabSize * EMU_PER_PX;
	}
	if (style.eaLineBreak !== undefined) out.eastAsianLineBreak = style.eaLineBreak;
	if (style.latinLineBreak !== undefined) out.latinLineBreak = style.latinLineBreak;
	if (style.fontAlignment !== undefined) out.fontAlign = style.fontAlignment;
	if (style.hangingPunctuation !== undefined) out.hangingPunctuation = style.hangingPunctuation;
	return out;
}

export function serializePlaceholderLevelStyle(
	style: PlaceholderTextLevelStyle,
	existing?: XmlObject,
): XmlObject {
	// Attribute values through the shared DrawingML codec (px -> EMU).
	const attrEdits: Record<string, string | null> = {};
	for (const [name, value] of drawingParagraphAttributeEntries(levelAttributesToDrawing(style))) {
		attrEdits[name] = value;
	}

	const childEdits = new Map<string, XmlObject | null>();
	if (style.lineSpacing !== undefined) {
		childEdits.set('a:lnSpc', objectTextSpacing({ unit: 'percent', value: style.lineSpacing }));
	} else if (style.lineSpacingExactPt !== undefined) {
		childEdits.set(
			'a:lnSpc',
			objectTextSpacing({ unit: 'points', value: style.lineSpacingExactPt }),
		);
	}
	if (style.spaceBefore !== undefined) {
		childEdits.set(
			'a:spcBef',
			serializeParagraphSpacing(
				style.spaceBefore,
				style.resolvedSpaceBefore,
				style.spaceBeforePercent,
			),
		);
	}
	if (style.spaceAfter !== undefined) {
		childEdits.set(
			'a:spcAft',
			serializeParagraphSpacing(
				style.spaceAfter,
				style.resolvedSpaceAfter,
				style.spaceAfterPercent,
			),
		);
	}
	applyBulletGroup(childEdits, style, existing?.['a:buFont'] as XmlObject | undefined);
	if (style.tabStops !== undefined) {
		childEdits.set(
			'a:tabLst',
			style.tabStops.length > 0 ? serializeTabStops(style.tabStops) : null,
		);
	}

	if (hasRunPropertyEdits(style)) {
		const existingDefRPr = existing?.['a:defRPr'] as XmlObject | undefined;
		childEdits.set('a:defRPr', serializeDefaultRunProperties(style, existingDefRPr));
	}

	return mergeOrderedXml(existing, attrEdits, childEdits, LEVEL_PPR_CHILD_ORDER);
}

/**
 * Serialize a resolved space-before/after value back to `a:spcBef` /
 * `a:spcAft`. An UNTOUCHED value (still equal to the snapshot taken at
 * parse time) that came from `a:spcPct` re-emits the original percentage,
 * so the spacing keeps scaling with a later font-size change instead of
 * being frozen to today's pixel value; anything else (no percent snapshot,
 * or the value has genuinely moved since parse) is written as `a:spcPts`,
 * matching what an editor working in absolute units produces.
 */
function serializeParagraphSpacing(
	value: number,
	resolvedValue: number | undefined,
	percent: number | undefined,
): XmlObject {
	if (percent !== undefined && value === resolvedValue) {
		return objectTextSpacing({ unit: 'percent', value: percent });
	}
	return objectTextSpacing({ unit: 'points', value: (value * PT_HUNDREDTHS_PER_PX) / 100 });
}

function hasRunPropertyEdits(style: PlaceholderTextLevelStyle): boolean {
	return (
		style.fontSize !== undefined ||
		style.bold !== undefined ||
		style.italic !== undefined ||
		style.color !== undefined ||
		style.fontFamily !== undefined
	);
}

function serializeDefaultRunProperties(
	style: PlaceholderTextLevelStyle,
	existing: XmlObject | undefined,
): XmlObject {
	const attrEdits: Record<string, string | null> = {};
	if (style.fontSize !== undefined) {
		attrEdits.sz = String(Math.round(style.fontSize * PT_HUNDREDTHS_PER_PX));
	}
	if (style.bold !== undefined) {
		attrEdits.b = style.bold ? '1' : '0';
	}
	if (style.italic !== undefined) {
		attrEdits.i = style.italic ? '1' : '0';
	}

	const childEdits = new Map<string, XmlObject | null>();
	if (style.color !== undefined) {
		// `colorChoiceXml` is the whole authored `a:solidFill`; keep it (theme
		// alias + transforms) unless the hex was edited away from it.
		const preserved = preservedIfStillCurrent(style.colorChoiceXml, style.color);
		childEdits.set('a:solidFill', preserved ?? srgbChoice(style.color));
		for (const variant of FILL_VARIANTS) {
			childEdits.set(variant, null);
		}
	}
	if (style.fontFamily !== undefined) {
		const latin: XmlObject = { ...((existing?.['a:latin'] as XmlObject | undefined) ?? {}) };
		// Untouched (still equal to what was parsed) and a raw alias was
		// captured: re-emit the alias (`+mj-lt`/`+mn-lt`) verbatim instead of
		// the literal name it resolves to today, or every save flattens the
		// theme reference. A genuine edit (the value has moved) always writes
		// the new literal name.
		latin['@_typeface'] =
			style.fontTypefaceXml !== undefined && style.fontFamily === style.resolvedFontFamily
				? style.fontTypefaceXml
				: style.fontFamily;
		childEdits.set('a:latin', latin);
	}

	return mergeOrderedXml(existing, attrEdits, childEdits, DEF_RPR_CHILD_ORDER);
}

function applyBulletGroup(
	childEdits: Map<string, XmlObject | null>,
	style: PlaceholderTextLevelStyle,
	existingBuFont: XmlObject | undefined,
): void {
	if (style.bulletColor !== undefined) {
		// A themed `a:schemeClr` bullet is re-emitted as authored; only a hex
		// edited away from the preserved node falls back to a literal srgb.
		const preserved = preservedIfStillCurrent(style.bulletColorXml, style.bulletColor);
		childEdits.set('a:buClr', preserved ?? srgbChoice(style.bulletColor));
	} else if (style.bulletColorXml !== undefined) {
		childEdits.set('a:buClr', style.bulletColorXml);
	}
	const size =
		style.bulletSizePercent !== undefined
			? ({ unit: 'percent', value: style.bulletSizePercent / 100 } as const)
			: style.bulletSizePts !== undefined
				? ({ unit: 'points', value: style.bulletSizePts } as const)
				: undefined;
	if (size) {
		const child = drawingBulletSizeChild(size);
		childEdits.set(`a:${child.local}`, objectAttributes(child.attributes));
		childEdits.set(size.unit === 'percent' ? 'a:buSzPts' : 'a:buSzPct', null);
	}
	if (style.bulletFontFamily !== undefined) {
		// Merge onto the EXISTING `a:buFont` node instead of replacing it, so
		// `panose`/`pitchFamily`/`charset` (fallback-face hints, not resolved
		// into the typed model) survive a save that only touches the typeface.
		childEdits.set('a:buFont', { ...existingBuFont, '@_typeface': style.bulletFontFamily });
	}

	const touchesType =
		style.bulletNone !== undefined ||
		style.bulletChar !== undefined ||
		style.bulletAutoNumType !== undefined;
	if (!touchesType) {
		return;
	}
	const typeNode = (bullet: DrawingBullet | undefined): XmlObject | null => {
		const child = bullet ? drawingBulletChild(bullet) : undefined;
		return child ? objectAttributes(child.attributes) : null;
	};
	childEdits.set('a:buNone', typeNode(style.bulletNone ? { kind: 'none' } : undefined));
	childEdits.set(
		'a:buAutoNum',
		typeNode(
			style.bulletAutoNumType !== undefined
				? { kind: 'autoNumber', scheme: style.bulletAutoNumType }
				: undefined,
		),
	);
	childEdits.set(
		'a:buChar',
		typeNode(style.bulletChar !== undefined ? { kind: 'char', char: style.bulletChar } : undefined),
	);
}

/** `a:tabLst` from typed tab stops (positions px -> EMU); mirrors the paragraph writer. */
function serializeTabStops(
	tabStops: NonNullable<PlaceholderTextLevelStyle['tabStops']>,
): XmlObject {
	return { 'a:tab': tabStops.map(serializeTabStop) };
}

function srgbChoice(hex: string): XmlObject {
	return { 'a:srgbClr': { '@_val': hex.replace('#', '') } };
}

/**
 * Decide whether a preserved colour node still describes `hex`.
 *
 * A themed choice (`a:schemeClr`, `a:sysClr`, ...) cannot be resolved here
 * (no theme in scope), so it is trusted and re-emitted as authored: an
 * editor that changes the colour is expected to clear the preserved node.
 * A literal `a:srgbClr` can be compared, so a hex that moved away from it
 * wins over the stale node.
 */
function preservedIfStillCurrent(
	preserved: XmlObject | undefined,
	hex: string,
): XmlObject | undefined {
	if (!preserved) {
		return undefined;
	}
	const srgb = preserved['a:srgbClr'] as XmlObject | undefined;
	if (srgb && !colorsEqual(String(srgb['@_val'] ?? ''), hex)) {
		return undefined;
	}
	return preserved;
}
