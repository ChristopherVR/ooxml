import { EMU_PER_PIXEL } from '../../../../units/index';
import { themeColorRefToSolidFillWithOpacity } from '../../color/theme-color-ref';
import { XmlObject, TextStyle } from '../../types';
import type { BulletInfo } from '../../types';
import { applyParagraphBulletOverrides } from '../../utils/paragraph-bullet-overrides';
import type { DrawingParagraphProperties } from '../../../../drawingml/text-model';
import {
	drawingBulletChild,
	drawingBulletSizeChild,
	drawingFontEntries,
	drawingParagraphAttributeEntries,
} from '../../../../drawingml/text-paragraph-codec';
import {
	assignObjectChild,
	objectAttributes,
	objectTextSpacing,
} from '../../utils/paragraph-properties-neutral';
import { serializeTabStop } from '../../utils/tab-stops';
import type { ParagraphChild } from './paragraph-child-assembly';
import { classifyParagraphChild, writeParagraphChildren } from './paragraph-child-assembly';

export { computeUniformSegmentOverrides } from './uniform-segment-overrides';

/** EMU-per-pixel conversion constant (matches PptxHandlerRuntime.EMU_PER_PX). */
export const EMU_PER_PX: number = EMU_PER_PIXEL;

/** Pre-computed spacing XML objects for the paragraph builder. */
export interface ParagraphSpacingConfig {
	spacingBefore: XmlObject | undefined;
	spacingAfter: XmlObject | undefined;
	lineSpacing: XmlObject | undefined;
	lineSpacingExactPt: number | undefined;
}

/**
 * Decide whether this paragraph may carry a given paragraph-scope property.
 *
 * `textStyle` reaching the builder is the SHAPE-level style merged with the
 * paragraph's own `a:pPr`, and the shape-level half is a *resolved* value: the
 * loader fills it, first-paragraph-wins, from the text body's `a:lstStyle`, the
 * inherited layout/master placeholder `a:pPr` and only then the paragraph's own
 * attributes. Writing all of it back onto every paragraph turned inherited
 * values into explicitly authored ones, which (a) churns every file on save and
 * (b) OVERRIDES the inheritance that would otherwise resolve per paragraph - so
 * a level-3 bullet was silently re-indented to the level-1 `marL` its shape
 * happened to resolve.
 *
 * Rule: emit a property the paragraph itself authored; otherwise emit the
 * shape-level value only when the paragraph authored NO properties of its own
 * AND sits at outline level 0, where the shape-level style is both the sole
 * description of it and was resolved through its own inheritance chain
 * (SDK-built decks, newly typed text, and any top-level `a:p` that carried no
 * `a:pPr`). A paragraph that authored an `a:pPr` keeps exactly the keys it
 * authored; everything else is left to inherit.
 *
 * The level test matters because the shape-level style is filled
 * first-paragraph-wins, so it describes level 0. Broadcasting it to a nested
 * bullet replaced that bullet's `a:lvl3pPr` indent with the level-1 one.
 */
function authoredPropertyGate(
	authoredProperties: TextStyle | undefined,
	level: number | undefined,
): (key: keyof TextStyle) => boolean {
	if (authoredProperties) {
		return (key) => authoredProperties[key] !== undefined;
	}
	const isNested = typeof level === 'number' && Number.isFinite(level) && level > 0;
	return () => !isNested;
}

const finitePx = (value: number | undefined): number | undefined =>
	typeof value === 'number' && Number.isFinite(value) ? value * EMU_PER_PX : undefined;

/**
 * The gated paragraph attributes of a {@link TextStyle} on the neutral model
 * (px -> EMU). `lvl` reaches here only when the paragraph carries one (its own
 * parsed `lvl`, including an authored `lvl="0"` PowerPoint itself often
 * writes, or one set by an edit), so it is written whenever present; a
 * paragraph that never had one gets none.
 */
function paragraphAttributesToDrawing(
	textStyle: TextStyle | undefined,
	paragraphAlign: string | undefined,
	level: number | undefined,
	owns: (key: keyof TextStyle) => boolean,
): DrawingParagraphProperties {
	const out: DrawingParagraphProperties = {};
	if (typeof level === 'number' && Number.isFinite(level) && level >= 0) {
		out.level = Math.min(Math.max(Math.round(level), 0), 8);
	}
	if (paragraphAlign && owns('align')) out.align = paragraphAlign;
	if (textStyle?.rtl !== undefined && owns('rtl')) out.rightToLeft = textStyle.rtl;
	const emu = (key: keyof TextStyle, value: number | undefined) =>
		owns(key) ? finitePx(value) : undefined;
	const marginLeftEmu = emu('paragraphMarginLeft', textStyle?.paragraphMarginLeft);
	if (marginLeftEmu !== undefined) out.marginLeftEmu = marginLeftEmu;
	const marginRightEmu = emu('paragraphMarginRight', textStyle?.paragraphMarginRight);
	if (marginRightEmu !== undefined) out.marginRightEmu = marginRightEmu;
	const indentEmu = emu('paragraphIndent', textStyle?.paragraphIndent);
	if (indentEmu !== undefined) out.indentEmu = indentEmu;
	const defaultTabSizeEmu = emu('defaultTabSize', textStyle?.defaultTabSize);
	if (defaultTabSizeEmu !== undefined) out.defaultTabSizeEmu = defaultTabSizeEmu;
	if (textStyle?.eaLineBreak !== undefined && owns('eaLineBreak')) {
		out.eastAsianLineBreak = textStyle.eaLineBreak;
	}
	if (textStyle?.latinLineBreak !== undefined && owns('latinLineBreak')) {
		out.latinLineBreak = textStyle.latinLineBreak;
	}
	if (textStyle?.fontAlignment && owns('fontAlignment')) out.fontAlign = textStyle.fontAlignment;
	if (textStyle?.hangingPunctuation !== undefined && owns('hangingPunctuation')) {
		out.hangingPunctuation = textStyle.hangingPunctuation;
	}
	return out;
}

/** Build the `a:pPr` (paragraph properties) XML object. */
export function buildParagraphPropertiesXml(
	textStyle: TextStyle | undefined,
	paragraphAlign: string | undefined,
	bulletInfo: BulletInfo | undefined,
	spacing: ParagraphSpacingConfig,
	level?: number,
	authoredProperties?: TextStyle,
): XmlObject {
	const owns = authoredPropertyGate(authoredProperties, level);
	// The attributes are written by the shared DrawingML codec, in
	// CT_TextParagraphProperties order (`marL marR lvl indent algn defTabSz
	// rtl eaLnBrk fontAlgn latinLnBrk hangingPunct`, as PowerPoint writes them).
	const paragraphProps = objectAttributes(
		drawingParagraphAttributeEntries(
			paragraphAttributesToDrawing(textStyle, paragraphAlign, level, owns),
		),
	);

	// CT_TextParagraphProperties child order (ECMA-376 21.1.2.2.7):
	//   lnSpc, spcBef, spcAft, <bullet group>, tabLst, defRPr, extLst.
	// The bullet group is itself ordered: buClr*, buSz*, buFont*, bu<type>.
	// fast-xml-parser serialises keys in insertion order, so assign in this
	// exact sequence — otherwise PowerPoint flags the file as corrupted.
	if (spacing.lineSpacing && owns('lineSpacing')) {
		paragraphProps['a:lnSpc'] = spacing.lineSpacing;
	} else if (
		typeof spacing.lineSpacingExactPt === 'number' &&
		Number.isFinite(spacing.lineSpacingExactPt) &&
		owns('lineSpacingExactPt')
	) {
		paragraphProps['a:lnSpc'] = objectTextSpacing({
			unit: 'points',
			value: spacing.lineSpacingExactPt,
		});
	}
	if (spacing.spacingBefore && owns('paragraphSpacingBefore')) {
		paragraphProps['a:spcBef'] = spacing.spacingBefore;
	}
	if (spacing.spacingAfter && owns('paragraphSpacingAfter')) {
		paragraphProps['a:spcAft'] = spacing.spacingAfter;
	}

	// Bullet properties. `resolveParagraphBulletInfo` walks the paragraph's
	// own `a:pPr`, the shape's `a:lstStyle`, the inherited placeholder and
	// the master's `a:defPPr` / `p:txStyles`, first-match-wins, so the
	// resolution can equally be the paragraph's OWN declaration or a purely
	// cascaded one. Writing a cascaded result unconditionally pins an
	// inherited bullet (a master `buFont="Arial"` / `buChar="•"`) onto every
	// paragraph the moment the slide is rewritten, including a body
	// placeholder paragraph with NO `a:pPr` at all: the "no authored
	// properties at all" exception every other per-paragraph field gets
	// (SDK-built decks, newly typed text) does not apply here, because for
	// those OTHER fields the resolved value is what the paragraph already
	// renders as either way, while a resolved BULLET (unlike align/margins)
	// is only "how the master looks right now" and must keep tracking it.
	// `ownedByParagraph` is `false` only when `resolveParagraphBulletInfo`
	// positively identified a cascaded source; it is `undefined` for bullet
	// info that never went through that function at all (an editor-added or
	// SDK-built bullet), which is written same as before this gate existed.
	if (bulletInfo && bulletInfo.ownedByParagraph !== false) {
		applyBulletProperties(paragraphProps, bulletInfo);
	} else {
		// The paragraph authored only a colour/size/font override for a bullet
		// it inherits: re-emit exactly that, read from the paragraph's OWN
		// properties (never the shape-level merge).
		applyParagraphBulletOverrides(paragraphProps, authoredProperties?.paragraphBulletPropertiesXml);
	}

	// Tab stops
	if (textStyle?.tabStops && textStyle.tabStops.length > 0 && owns('tabStops')) {
		paragraphProps['a:tabLst'] = { 'a:tab': textStyle.tabStops.map(serializeTabStop) };
	} else if (authoredProperties?.tabStopsExplicitEmpty) {
		// This specific paragraph authored `<a:tabLst/>` with no tab stops
		// (see `tabStopsExplicitEmpty`): re-emit it exactly, regardless of
		// what `textStyle.tabStops` resolved to from elsewhere in the cascade.
		paragraphProps['a:tabLst'] = {};
	}

	// `a:defRPr` is the paragraph default run properties. It follows the bullet group
	// and `a:tabLst`, and precedes `a:extLst`.
	if (textStyle?.paragraphDefaultRunPropertiesXml && owns('paragraphDefaultRunPropertiesXml')) {
		paragraphProps['a:defRPr'] = textStyle.paragraphDefaultRunPropertiesXml;
	}

	// `a:extLst` is the very last child of CT_TextParagraphProperties. Re-emit
	// the captured opaque subtree verbatim when present so authored extensions
	// survive a round-trip.
	if (textStyle?.paragraphPropertiesExtLstXml && owns('paragraphPropertiesExtLstXml')) {
		paragraphProps['a:extLst'] = textStyle.paragraphPropertiesExtLstXml;
	}

	return paragraphProps;
}

/** The bullet size of a {@link BulletInfo} on the neutral model (`text` wins, then percent). */
function bulletSizeToDrawing(
	bulletInfo: BulletInfo,
): DrawingParagraphProperties['bulletSize'] | undefined {
	if (bulletInfo.sizeInherit) return 'text';
	if (bulletInfo.sizePercent !== undefined) {
		return { unit: 'percent', value: bulletInfo.sizePercent / 100 };
	}
	if (bulletInfo.sizePts !== undefined) return { unit: 'points', value: bulletInfo.sizePts };
	return undefined;
}

/** Apply bullet-related XML attributes from {@link BulletInfo} into `paragraphProps`. */
export function applyBulletProperties(paragraphProps: XmlObject, bulletInfo: BulletInfo): void {
	// CT_TextParagraphProperties bullet-group schema order:
	//   buClr (color), buSzPct/buSzPts (size), buFont (typeface),
	//   buNone/buAutoNum/buChar/buBlip (type). fast-xml-parser serialises
	//   keys in insertion order, so assign in this exact sequence or
	//   PowerPoint's validator rejects the run.
	//
	// The color/size/font "inherit from text" markers (`buClrTx`/`buSzTx`/
	// `buFontTx`) are each their OWN choice group (EG_TextBulletColor/
	// Size/Typeface) and are independent of the bullet TYPE choice
	// (`buNone`/`buAutoNum`/`buChar`/`buBlip`, EG_TextBulletType) — a
	// paragraph can author `buNone` and still say "inherit bullet colour
	// from text" for when a bullet mark IS shown again later in the
	// cascade. Emitting `buNone` used to `return` immediately, before these
	// three groups ran, silently dropping them on every round-trip.
	//
	// Inherit-from-text variants take precedence over the explicit
	// `buClr` / `buSzPct|Pts` / `buFont` declarations: when both forms are
	// present the schema only allows one. Emit `<a:buClrTx/>` etc. when the
	// parsed model captured the marker.
	if (bulletInfo.colorInherit) {
		paragraphProps['a:buClrTx'] = {};
	} else if (bulletInfo.colorRef) {
		// A typed theme ref wins: keeps the bullet following the theme palette
		// after a later theme change instead of freezing today's sRGB/schemeClr.
		paragraphProps['a:buClr'] = themeColorRefToSolidFillWithOpacity(bulletInfo.colorRef);
	} else if (bulletInfo.colorXml) {
		// Re-emit the original colour-choice node (a:schemeClr / a:sysClr /
		// a:prstClr / a:srgbClr plus any colour transforms) verbatim so themed
		// bullet colours survive a round-trip.
		paragraphProps['a:buClr'] = bulletInfo.colorXml;
	} else if (bulletInfo.color) {
		const colorHex = bulletInfo.color.replace('#', '');
		paragraphProps['a:buClr'] = {
			'a:srgbClr': { '@_val': colorHex },
		};
	}
	// Size, font and type children through the shared DrawingML codec.
	const size = bulletSizeToDrawing(bulletInfo);
	if (size) {
		assignObjectChild(paragraphProps, drawingBulletSizeChild(size));
	}
	if (bulletInfo.fontInherit) {
		paragraphProps['a:buFontTx'] = {};
	} else if (bulletInfo.fontFamily) {
		// `a:buFont` is a CT_TextFont, same as `a:latin`/`a:ea`/`a:cs`/`a:sym`:
		// re-emit the font-matching hints those already carry, or a bullet's
		// own PANOSE/pitch-family/charset silently vanish on every save.
		paragraphProps['a:buFont'] = objectAttributes(
			drawingFontEntries({
				typeface: bulletInfo.fontFamily,
				...(bulletInfo.fontPanose ? { panose: bulletInfo.fontPanose } : {}),
				...(bulletInfo.fontPitchFamily !== undefined
					? { pitchFamily: bulletInfo.fontPitchFamily }
					: {}),
				...(bulletInfo.fontCharset !== undefined ? { charset: bulletInfo.fontCharset } : {}),
			}),
		);
	}
	// Bullet TYPE choice (EG_TextBulletType): exactly one of none/char/
	// autoNum/blip. `none` is checked first and returns, since `bulletInfo`
	// never carries more than one of these at once.
	if (bulletInfo.none) {
		assignObjectChild(paragraphProps, drawingBulletChild({ kind: 'none' })!);
		return;
	}
	if (bulletInfo.char) {
		assignObjectChild(paragraphProps, drawingBulletChild({ kind: 'char', char: bulletInfo.char })!);
	}
	if (bulletInfo.autoNumType) {
		const startAt =
			bulletInfo.autoNumStartAt !== undefined && bulletInfo.autoNumStartAt !== 1
				? { startAt: bulletInfo.autoNumStartAt }
				: {};
		assignObjectChild(
			paragraphProps,
			drawingBulletChild({ kind: 'autoNumber', scheme: bulletInfo.autoNumType, ...startAt })!,
		);
	}
	if (bulletInfo.imageBlipFillXml) {
		// Re-emit the captured `a:buBlip` subtree verbatim (a:blip + a:extLst,
		// a:tile, a:stretch, a:srcRect) so picture-bullet modifiers such as a
		// crop or tile setting survive a round-trip, rather than reconstructing
		// a bare `a:blip[@r:embed]` that drops every modifier.
		paragraphProps['a:buBlip'] = bulletInfo.imageBlipFillXml;
	} else if (bulletInfo.imageRelId) {
		paragraphProps['a:buBlip'] = {
			'a:blip': { '@_r:embed': bulletInfo.imageRelId },
		};
	}
}

/** Markup-only facts about the source paragraph; see `paragraph-markup-flags.ts`. */
export interface ParagraphMarkupFlags {
	/** The source wrote an empty `<a:pPr/>`. */
	emptyParagraphPropertiesAuthored?: boolean | undefined;
	/** The source had no run content and no `a:endParaRPr`. */
	bareParagraph?: boolean | undefined;
}

/** Assemble a paragraph XML object from runs and pre-built paragraph properties. */
export function assembleParagraphXml(
	runs: XmlObject[],
	paragraphProps: XmlObject,
	endParaRunProperties?: Record<string, unknown>,
	markup?: ParagraphMarkupFlags,
): XmlObject {
	// OOXML CT_TextParagraph requires child order: pPr?, (r|br|fld)*, endParaRPr?.
	// Since fast-xml-parser serialises keys in insertion order, build the
	// object in that exact sequence.
	//
	// `a:pPr` is OPTIONAL (CT_TextParagraph, ECMA-376 21.1.2.2.7): a paragraph
	// that authored no properties of its own parses `<a:pPr/>` and "no pPr at
	// all" to the SAME empty result (fast-xml-parser gives a childless,
	// attribute-less element back as `''`, which is falsy exactly like
	// `undefined`), so `buildParagraphPropertiesXml` returns `{}` for both.
	// Writing that empty object back as a key materialized `<a:pPr></a:pPr>`
	// on every paragraph that never had one, purely because the key was always
	// assigned regardless of whether it carried anything. Both sides of that
	// collapsed distinction re-emit as "no pPr", which is lossless: an empty
	// `<a:pPr/>` carries no attributes or children, so dropping it and never
	// having had one are semantically identical. The parser records which of
	// the two the source wrote (`emptyParagraphPropertiesAuthored`), so a
	// PowerPoint-authored `<a:pPr/>` is still given back as written.
	const paragraph: XmlObject = {};
	if (Object.keys(paragraphProps).length > 0 || markup?.emptyParagraphPropertiesAuthored) {
		paragraph['a:pPr'] = paragraphProps;
	}

	// `runs` already arrives in segment order, so the authored sequence of
	// runs / fields / breaks / inline math is simply its order.
	const children = runs
		.map((run) => classifyParagraphChild(run))
		.filter((child): child is ParagraphChild => child !== undefined);

	if (children.length > 0) {
		writeParagraphChildren(paragraph, children);
	} else {
		// Every run was an equation marker with no captured XML (or there were
		// no runs at all): fall back to emitting whatever was handed in.
		paragraph['a:r'] = runs.length > 1 ? runs : runs[0];
	}

	// Re-emit parsed end-paragraph run properties verbatim. When none were
	// captured, fall back to the minimal `lang="en-US"` stub PowerPoint itself
	// emits for new paragraphs, but ONLY for a paragraph with no run content:
	// that stub is what a blank line's style lives in (SDK-built or authored),
	// so a genuinely empty paragraph still needs something to carry it. A
	// paragraph that already has real run content but happened to author no
	// `a:endParaRPr` of its own is untouched content, not a blank line, and
	// must pass through without inventing a trailing endParaRPr its source
	// never had (measured: fabricated `<a:endParaRPr lang="en-US"/>` after the
	// last authored run of every such paragraph in a rewritten slide). A
	// source paragraph that was already runless WITHOUT one (`bareParagraph`,
	// a bare `<a:p/>`) goes back the same way.
	if (endParaRunProperties && typeof endParaRunProperties === 'object') {
		paragraph['a:endParaRPr'] = endParaRunProperties as XmlObject;
	} else if (runs.length === 0 && !markup?.bareParagraph) {
		paragraph['a:endParaRPr'] = { '@_lang': 'en-US' };
	}

	return paragraph;
}
