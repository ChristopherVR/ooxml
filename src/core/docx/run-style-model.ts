import type { StyleMetadata } from './style-metadata';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model';

/** Direct run properties supported by the editor for style/default resolution, in native Word units. */
export type RunFormatting = Pick<
	TextRun,
	| 'bold'
	| 'italic'
	| 'underline'
	| 'underlineStyle'
	| 'underlineColor'
	| 'strike'
	| 'doubleStrike'
	| 'caps'
	| 'smallCaps'
	| 'vanish'
	| 'language'
	| 'highlight'
	| 'verticalAlign'
	| 'fontSize'
	| 'fontFamily'
	| 'fontFamilyEastAsia'
	| 'fontFamilyComplexScript'
	| 'fontSizeComplexScript'
	| 'boldComplexScript'
	| 'italicComplexScript'
	| 'fontTheme'
	| 'color'
	| 'colorTheme'
	| 'characterSpacingTwips'
	| 'textScalePercent'
	| 'ligatures'
	| 'kerningHalfPoints'
	| 'positionHalfPoints'
	| 'shadingFill'
	| 'shadingThemeFill'
>;
/** A `styles.xml` style carrying run formatting: a character style, or a paragraph style's `w:rPr`. */
export interface CharacterStyleDefinition extends StyleMetadata {
	type: 'character' | 'paragraph';
	/** `w:link`; the paired style of the opposite kind sharing this style's look. */
	linkedStyle?: string;
	formatting: RunFormatting;
	/** `w:rPr/w:lang/@w:eastAsia`, which selects the kinsoku set of the style's text. */
	eastAsiaLanguage?: string;
}
/** Parsed run-level style data (docDefaults `rPrDefault` plus character/paragraph style `rPr`). */
export interface RunStyleCatalog {
	docDefaults: RunFormatting;
	/** `w:rPrDefault`'s `w:lang/@w:eastAsia`. */
	docDefaultsEastAsiaLanguage?: string;
	styles: Record<string, CharacterStyleDefinition>;
	/** Ids of styles kept as source XML, such as numbering styles. */
	reservedStyleIds?: string[];
	warnings: string[];
}
