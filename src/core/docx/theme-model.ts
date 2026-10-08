// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { ThemeColorSlot } from '../drawingml/theme-model';
import type { StThemeColor } from './generated/wml-simple-types';
/** Theme color scheme slots (`a:clrScheme` children): the neutral DrawingML slots. */
export type { ThemeColorSlot };
/** WordprocessingML `ST_ThemeColor` tokens (generated from the schema) referenced from `w:themeColor`/`w:themeFill`. */
export type ThemeColorToken = StThemeColor;
export interface ThemeFontSet {
	latin?: string;
	eastAsia?: string;
	complexScript?: string;
}
/** Parsed word/theme/theme1.xml plus word/settings.xml clrSchemeMapping; never flattened onto runs. */
export interface ThemeCatalog {
	colors: Partial<Record<ThemeColorSlot, string>>;
	colorMapping: Partial<Record<'bg1' | 'tx1' | 'bg2' | 'tx2', ThemeColorSlot>>;
	fonts: { major: ThemeFontSet; minor: ThemeFontSet };
}
/** A direct theme color reference kept alongside (not instead of) any resolved RGB fallback. */
export interface ThemeColorReference {
	token: ThemeColorToken;
	/** Fraction in [0,1], parsed from the two-digit hex byte Word stores. */
	tint?: number;
	shade?: number;
}
export type ThemeFontScript = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';
export type ThemeFontRole = 'major' | 'minor';
/** Word `ST_Underline` styles beyond the plain single-line toggle. */
export type WordUnderlineStyle =
	| 'single'
	| 'words'
	| 'double'
	| 'thick'
	| 'dotted'
	| 'dottedHeavy'
	| 'dash'
	| 'dashedHeavy'
	| 'dashLong'
	| 'dashLongHeavy'
	| 'dotDash'
	| 'dashDotHeavy'
	| 'dotDotDash'
	| 'dashDotDotHeavy'
	| 'wave'
	| 'wavyHeavy'
	| 'wavyDouble'
	| 'none';
