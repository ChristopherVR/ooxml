// Format-neutral DrawingML rich text model (ECMA-376 Part 1, 21.1.2): body properties
// (`a:bodyPr`), list styles (`a:lstStyle`), paragraph properties (`a:pPr`, `a:lvlNpPr`), run
// properties (`a:rPr`, `a:defRPr`, `a:endParaRPr`), runs, line breaks and fields. Values keep the
// file's units (EMU, points, degrees, fractions) so a reader in any product can convert them, and
// children the model does not cover are kept as XML so a write puts them back.
//
// The small `DrawingTextBody` in `types.ts` is the flattened summary charts use; this is the full
// model SmartArt and shape text build on.
import type { DrawingColor, DrawingFill, DrawingLine, DrawingTextSpacing } from './types';

/** A `CT_TextFont` slot (`a:latin`, `a:ea`, `a:cs`, `a:sym`, `a:buFont`). */
export interface DrawingFont {
	typeface: string;
	panose?: string;
	pitchFamily?: number;
	charset?: number;
}

/** `a:hlinkClick` / `a:hlinkMouseOver` (CT_Hyperlink); the target stays a relationship id. */
export interface DrawingHyperlink {
	relId?: string;
	tooltip?: string;
	action?: string;
	invalidUrl?: string;
	targetFrame?: string;
	history?: boolean;
	highlightClick?: boolean;
	endSound?: boolean;
	/** Children such as `a:snd` and `a:extLst`, kept verbatim. */
	extraXml?: string[];
}

/** Run properties (`CT_TextCharacterProperties`). */
export interface DrawingRunProperties {
	/** `sz` in points (the file stores hundredths). */
	sizePt?: number;
	bold?: boolean;
	italic?: boolean;
	/** `u`: `sng`, `dbl`, `none`, ... as written. */
	underline?: string;
	/** `strike`: `noStrike`, `sngStrike`, `dblStrike`. */
	strike?: string;
	/** `cap`: `none`, `small`, `all`. */
	caps?: string;
	/** `baseline` as a fraction of the font size (0.3 = 30% superscript). */
	baseline?: number;
	/** `spc` character spacing in points. */
	spacingPt?: number;
	/** `kern` minimum font size for kerning, in points. */
	kernPt?: number;
	language?: string;
	altLanguage?: string;
	kumimoji?: boolean;
	normalizeHeight?: boolean;
	noProof?: boolean;
	dirty?: boolean;
	spellingError?: boolean;
	smartTagClean?: boolean;
	smartTagId?: number;
	bookmark?: string;
	/** Text outline (`a:ln`); `lineSourceXml` keeps join, compound and arrow children. */
	line?: DrawingLine;
	lineSourceXml?: string;
	/** Text fill (`a:solidFill`, `a:gradFill`, `a:noFill`...). */
	fill?: DrawingFill;
	highlight?: DrawingColor;
	/** `a:uLnTx` (true) or nothing; an explicit underline line is kept in `underlineLineXml`. */
	underlineLineFollowsText?: boolean;
	underlineLineXml?: string;
	/** `a:uFillTx` (true) or nothing; an explicit underline fill is `underlineFill`. */
	underlineFillFollowsText?: boolean;
	underlineFill?: DrawingFill;
	latin?: DrawingFont;
	eastAsian?: DrawingFont;
	complexScript?: DrawingFont;
	symbol?: DrawingFont;
	hyperlinkClick?: DrawingHyperlink;
	hyperlinkMouseOver?: DrawingHyperlink;
	/** `a:rtl` child element (not an attribute on runs). */
	rightToLeft?: boolean;
	/** Children the model does not cover (`a:effectLst`, `a:extLst`...), verbatim. */
	extraXml?: string[];
}

/** One `a:tab` of a `a:tabLst`. */
export interface DrawingTabStop {
	/** Position in EMU. */
	positionEmu?: number;
	/** `l`, `ctr`, `r`, `dec`. */
	align?: string;
}

/** The bullet of a paragraph (`a:buNone`, `a:buChar`, `a:buAutoNum`, `a:buBlip`). */
export type DrawingBullet =
	| { kind: 'none' }
	| { kind: 'char'; char: string }
	| { kind: 'autoNumber'; scheme: string; startAt?: number }
	| { kind: 'picture'; xml: string };

/** Paragraph properties (`CT_TextParagraphProperties`). */
export interface DrawingParagraphProperties {
	/** `lvl` (0-8). */
	level?: number;
	/** `algn`: `l`, `ctr`, `r`, `just`, `justLow`, `dist`, `thaiDist`. */
	align?: string;
	/** `marL`, `marR` and `indent` in EMU. */
	marginLeftEmu?: number;
	marginRightEmu?: number;
	indentEmu?: number;
	/** `defTabSz` in EMU. */
	defaultTabSizeEmu?: number;
	rightToLeft?: boolean;
	eastAsianLineBreak?: boolean;
	latinLineBreak?: boolean;
	hangingPunctuation?: boolean;
	/** `fontAlgn`: `auto`, `t`, `ctr`, `base`, `b`. */
	fontAlign?: string;
	lineSpacing?: DrawingTextSpacing;
	spaceBefore?: DrawingTextSpacing;
	spaceAfter?: DrawingTextSpacing;
	/** `a:buClrTx` is `'text'`. */
	bulletColor?: DrawingColor | 'text';
	/** `a:buSzTx` is `'text'`; `a:buSzPct` a fraction; `a:buSzPts` points. */
	bulletSize?: 'text' | { unit: 'percent' | 'points'; value: number };
	/** `a:buFontTx` is `'text'`. */
	bulletFont?: DrawingFont | 'text';
	bullet?: DrawingBullet;
	tabStops?: DrawingTabStop[];
	defaultRunProperties?: DrawingRunProperties;
	extraXml?: string[];
}

/** `a:bodyPr` autofit choice. */
export type DrawingAutofit =
	| { kind: 'none' }
	| { kind: 'shape' }
	| { kind: 'normal'; fontScale?: number; lineSpacingReduction?: number };

/** Body properties (`CT_TextBodyProperties`). */
export interface DrawingBodyProperties {
	/** `rot` in degrees. */
	rotation?: number;
	spaceFirstLastParagraph?: boolean;
	verticalOverflow?: string;
	horizontalOverflow?: string;
	/** `vert`: `horz`, `vert`, `vert270`, `eaVert`, ... */
	vertical?: string;
	/** `wrap`: `none` or `square`. */
	wrap?: string;
	/** Insets in EMU. */
	insetLeftEmu?: number;
	insetTopEmu?: number;
	insetRightEmu?: number;
	insetBottomEmu?: number;
	columnCount?: number;
	columnSpacingEmu?: number;
	rightToLeftColumns?: boolean;
	fromWordArt?: boolean;
	/** `anchor`: `t`, `ctr`, `b`, `just`, `dist`. */
	anchor?: string;
	anchorCenter?: boolean;
	forceAntiAlias?: boolean;
	upright?: boolean;
	compatibleLineSpacing?: boolean;
	autofit?: DrawingAutofit;
	scene3d?: import('./text-3d-model').DrawingTextScene3D;
	shape3d?: import('./text-3d-model').DrawingTextShape3D;
	flatText3d?: import('./text-3d-model').DrawingFlatText3D;
	/** `a:prstTxWarp`, `a:scene3d`, `a:sp3d`, `a:flatTx`, `a:extLst`, verbatim. */
	extraXml?: string[];
}

/** An ordered child of a paragraph. */
export type DrawingParagraphItem =
	| { kind: 'run'; text: string; properties?: DrawingRunProperties }
	| { kind: 'break'; properties?: DrawingRunProperties }
	| {
			kind: 'field';
			id: string;
			type?: string;
			text: string;
			properties?: DrawingRunProperties;
			paragraphProperties?: DrawingParagraphProperties;
	  }
	/** A child the model does not cover (`a14:m` equations, `mc:AlternateContent`...), verbatim. */
	| { kind: 'unknown'; xml: string };

/** One `a:p`. */
export interface DrawingParagraph {
	properties?: DrawingParagraphProperties;
	items: DrawingParagraphItem[];
	/** `a:endParaRPr`. */
	endProperties?: DrawingRunProperties;
}

/** A list style (`a:lstStyle`): the default and nine level paragraph properties. */
export interface DrawingListStyle {
	default?: DrawingParagraphProperties;
	/** Index 0 is `a:lvl1pPr`. Missing levels are `undefined`. */
	levels: (DrawingParagraphProperties | undefined)[];
	extraXml?: string[];
}

/** A full DrawingML text body (`a:txBody`, `dsp:txBody`, `dgm:t`, `c:rich`...). */
export interface DrawingRichTextBody {
	bodyProperties?: DrawingBodyProperties;
	listStyle?: DrawingListStyle;
	paragraphs: DrawingParagraph[];
}
