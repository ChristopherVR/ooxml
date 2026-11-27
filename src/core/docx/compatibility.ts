// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// One typed view of the layout compatibility settings in `word/settings.xml` (`w:compat`), and
// named predicates for every layout rule that depends on them, so the layout engine never compares
// raw mode numbers. Word writes `compatibilityMode` ([MS-DOCX]) as 11 (Word 2003), 12 (2007),
// 14 (2010) or 15 (2013 and later); a document without one is treated as legacy.
import { children, first, getW, type XmlElement } from './xml';
import { onOffElement, parseOnOff } from './simple-types';

export const WORD_COMPAT_URI = 'http://schemas.microsoft.com/office/word';

/** The `w:compat` values the document model keeps. */
export interface CompatibilitySettings {
	/** `w:compatSetting[@w:name="compatibilityMode"]`, a positive integer. */
	compatibilityMode?: number;
	overrideTableStyleFontSizeAndJustification?: boolean;
	/** `w:compat/w:doNotUseHTMLParagraphAutoSpacing`. */
	doNotUseHtmlParagraphAutoSpacing?: boolean;
	doNotExpandShiftReturn?: boolean;
}

/** `modern` is compatibility mode 15 and later; `legacy` is 11 to 14 and an absent mode. */
export type CompatibilityGeneration = 'modern' | 'legacy';

export interface CompatibilityProfile {
	generation: CompatibilityGeneration;
	/** The saved mode number, when the document has one. */
	mode?: number;
	/** Legacy `w:compat` child options. */
	doNotUseHtmlParagraphAutoSpacing: boolean;
	doNotExpandShiftReturn: boolean;
}

/** Reads the compatibility mode and legacy options from `w:settings/w:compat`. */
export function readCompatibilitySettings(compat: XmlElement | undefined): CompatibilitySettings {
	if (!compat) return {};
	const settings = children(compat, 'compatSetting');
	const value = (name: string) =>
		getW(
			settings.find((item) => getW(item, 'name') === name && getW(item, 'uri') === WORD_COMPAT_URI),
			'val',
		);
	const mode = Number(value('compatibilityMode'));
	const override = parseOnOff(value('overrideTableStyleFontSizeAndJustification'));
	return {
		...(Number.isInteger(mode) && mode > 0 ? { compatibilityMode: mode } : {}),
		...(override === undefined ? {} : { overrideTableStyleFontSizeAndJustification: override }),
		...(onOffElement(first(compat, 'doNotUseHTMLParagraphAutoSpacing')) === true
			? { doNotUseHtmlParagraphAutoSpacing: true }
			: {}),
		...(onOffElement(first(compat, 'doNotExpandShiftReturn')) === true
			? { doNotExpandShiftReturn: true }
			: {}),
	};
}

/** The layout profile of a document model (or of the settings read from its package). */
export function compatibilityProfile(settings: CompatibilitySettings): CompatibilityProfile {
	const mode = settings.compatibilityMode;
	return {
		generation: mode !== undefined && mode >= 15 ? 'modern' : 'legacy',
		...(mode !== undefined ? { mode } : {}),
		doNotUseHtmlParagraphAutoSpacing: settings.doNotUseHtmlParagraphAutoSpacing === true,
		doNotExpandShiftReturn: settings.doNotExpandShiftReturn === true,
	};
}

/**
 * `doNotUseHTMLParagraphAutoSpacing`: adjacent paragraph spacing adds the space after and before
 * instead of keeping the larger, and auto spacing reads the twips. Measured in modes 12, 14 and 15.
 */
export function addsParagraphSpacing(profile: CompatibilityProfile): boolean {
	return profile.doNotUseHtmlParagraphAutoSpacing;
}

/** Word's modern justification can squeeze word spaces to fit an extra word. */
export function squeezesJustifiedSpaces(profile: CompatibilityProfile): boolean {
	return profile.generation === 'modern';
}

/** Word mode 15 moves an anchor with its wrapping float when the float exceeds the text area. */
export function keepsWrappingFloatsAboveBottomMargin(profile: CompatibilityProfile): boolean {
	return profile.mode === 15;
}

/**
 * A `pageBreakBefore` paragraph at the top of a page keeps its space before. Measured: kept in
 * modes 12 and 14, dropped in 15; other modes (including later numbers) keep the legacy rule
 * until they are measured.
 */
export function keepsSpaceBeforePageBreak(profile: CompatibilityProfile): boolean {
	return profile.mode !== 15;
}

/**
 * Native mode 15 footnote layout: the separator repeats the normal story on continued pages and
 * `w:footnotePr/w:numCols` gives the notes their own columns. Only mode 15 has been measured.
 */
export function usesModernFootnoteLayout(profile: CompatibilityProfile): boolean {
	return profile.mode === 15;
}

/** Modes 12 and 14 (measured) ignore the footnote column count and attach notes to body columns. */
export function attachesFootnotesToBodyColumns(profile: CompatibilityProfile): boolean {
	return profile.mode === 12 || profile.mode === 14;
}

/**
 * AutoFit tables size against the text column plus the first cell's left margin, and a percentage
 * width against the column plus both outer cell margins. Measured in modes 12, 14 and with no
 * mode (table-autofit fixtures); mode 15 sizes against the column itself.
 */
export function autofitOverhangsCellMargins(profile: CompatibilityProfile): boolean {
	return profile.generation === 'legacy';
}

/**
 * `layoutInCell="0"` puts a floating object anchored in a table cell on the page; from mode 15
 * Word lays every such object out in its cell. Measured in modes 14 and 15 (cell-floats fixtures).
 */
export function honoursLayoutInCell(profile: CompatibilityProfile): boolean {
	return profile.generation === 'legacy';
}
