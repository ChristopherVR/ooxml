/**
 * Engine-side input contracts. These are intentionally a separate type family
 * from `@christophervr/docx-core`'s `DocumentModel`: several fields here
 * (keepNext, cantSplit, tblHeader, explicit page/column breaks, sections,
 * columns) are not yet represented in the shared model. `adapter.ts` maps
 * what the current model exposes and leaves the rest at their (Word-default)
 * fallback, so the pipeline keeps compiling as other agents add model
 * fields, and nothing here ever reads `(x as any)`.
 */
import type { SignedTwips, TextRun, Twips } from '../index';

export interface LayoutRun {
	text: string;
	/** Footnote referenced by this run; its note starts on the page containing this marker. */
	footnoteReferenceId?: string;
	/** Omit from layout while retaining its text length in source ranges for following runs. */
	hidden?: boolean;
	/** Source length when displayed text is absent or generated; otherwise defaults to text.length. */
	sourceLength?: number;
	bold?: boolean;
	italic?: boolean;
	fontFamily?: string;
	/** Points, matching `TextRun.fontSize` in docx-core. */
	fontSizePt?: number;
	textScalePercent?: number;
	ligatures?: import('../ligatures').Ligatures;
	characterSpacingPx?: number;
	kerningThresholdPt?: number;
	/** Positive raises the baseline, negative lowers it. */
	positionPx?: number;
	/** A hard page/column break placed immediately after this run's text. */
	breakAfter?: 'page' | 'column';
	/** An inline picture: occupies its size on the line (`text` is empty). */
	object?: LayoutObject;
	/** Generated text (a list label) that is not part of the paragraph's own characters. */
	synthetic?: boolean;
	/** Generated list marker; the remaining characters in this run are its separator. */
	marker?: {
		length: number;
		alignment: 'left' | 'center' | 'right';
		/** Picture bullet: the marker occupies this width instead of its fallback glyph's. */
		widthPx?: number;
	};
	/** Superscript or subscript: drawn smaller and raised or lowered (note marks are superscript). */
	script?: 'super' | 'sub';
	/** `#rrggbb`, display only. */
	color?: string;
	/** `#rrggbb` shading behind the run's text, display only (Word's form field shading). */
	background?: string;
	/** `w:ptab`: the run's tabs go to a fixed point of the line, whatever the tab stops. */
	positionalTab?: TextRun['positionalTab'];
	underline?: boolean;
	strike?: boolean;
	/**
	 * The East Asian language whose kinsoku set governs this run's characters (from the resolved
	 * `w:lang/@w:eastAsia`); absent, the run's characters are never kept off a line start or end.
	 */
	kinsokuLanguage?: import('../../text/kinsoku').KinsokuLanguage;
}

/** A picture's package part and size in CSS pixels; the engine never loads its bytes. */
export interface LayoutObject {
	pictureFill?: Extract<import('../../drawingml/types').DrawingFill, { kind: 'picture' }>;
	partName: string;
	contentType: string;
	widthPx: number;
	heightPx: number;
	/** How far the object extends below the baseline (an equation's descent); 0 sits on it. */
	descentPx?: number;
	/**
	 * A chart drawn at this size: the painted SVG markup, or only its frame label.
	 * `shape` marks a drawing shape painted the same way, whose outline may reach
	 * past its box and which has no chart background; `textColor` is the colour of its text box
	 * paragraphs where their runs set none.
	 */
	chart?: { svg?: string; label: string; shape?: true; textColor?: string };
	/** A shape's text box content, drawn over its shape by the print view; the engine never lays it out. */
	textBox?: import('../inline-model').TextBoxContent;
	/** Native note separator rule, drawn without loading a media part. */
	noteSeparator?: 'separator' | 'continuationSeparator';
}

/** A floating picture anchored to a paragraph (`wp:anchor`), positioned on the page. */
export interface LayoutFloat extends LayoutObject {
	/** Which of the paragraph's floating pictures this is, in run order. */
	index?: number;
	/** `wp:positionH/@relativeFrom`: page, margin, column, leftMargin, rightMargin, character… */
	relativeFromH?: string;
	alignH?: 'left' | 'center' | 'right' | 'inside' | 'outside';
	offsetXPx?: number;
	/** `wp:positionV/@relativeFrom`: page, margin, paragraph, line, topMargin, bottomMargin… */
	relativeFromV?: string;
	alignV?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
	offsetYPx?: number;
	behindText?: boolean;
	wrap?: string;
	/** `wrapText` of a square, tight or through wrap (bothSides, left, right, largest). */
	wrapText?: string;
	/** Where the anchor sits in the paragraph's source text, for `line` and `character` frames. */
	sourceOffset?: number;
	/**
	 * Anchored in a table cell but positioned on the page (`layoutInCell="0"`, honoured before
	 * compatibility mode 15). Otherwise a float in a cell is laid out in it (`cell-floats.ts`).
	 */
	layoutInCell?: false;
}

/** A paragraph tab stop, positioned from the paragraph's text margin (not its indent). */
export interface LayoutTabStop {
	posPx: number;
	align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'start' | 'end' | 'num';
	leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

/** A paragraph border side: the line and its `w:space` gap to the text. */
export interface LayoutParagraphBorder extends LayoutBorder {
	spacePx: number;
}
export interface LayoutParagraphBorders {
	top?: LayoutParagraphBorder;
	bottom?: LayoutParagraphBorder;
	left?: LayoutParagraphBorder;
	right?: LayoutParagraphBorder;
}

/** A footnote's content, laid out in the page's footnote area. */
export interface LayoutNote {
	id: string;
	paragraphs: LayoutParagraph[];
}

export type ParagraphAlign = 'left' | 'center' | 'right' | 'justify';
export type LineSpacingRule = 'auto' | 'exact' | 'atLeast';

export interface LayoutParagraph {
	/** Mandatory paragraph after a table, outside the preceding table's column balance. */
	afterTableSectionBreak?: boolean;
	kind: 'paragraph';
	id: string;
	runs: LayoutRun[];
	align?: ParagraphAlign;
	direction?: 'ltr' | 'rtl';
	spacingBeforeTwips?: Twips;
	spacingAfterTwips?: Twips;
	lineSpacingTwips?: SignedTwips;
	lineSpacingRule?: LineSpacingRule;
	indentLeftTwips?: SignedTwips;
	indentRightTwips?: SignedTwips;
	indentStartTwips?: SignedTwips;
	indentEndTwips?: SignedTwips;
	firstLineTwips?: Twips;
	hangingTwips?: Twips;
	/** `w:contextualSpacing`: suppress spacing between paragraphs sharing `styleId`. */
	contextualSpacing?: boolean;
	styleId?: string;
	/** `w:pageBreakBefore`. */
	pageBreakBefore?: boolean;
	/** `w:keepNext`: keep this paragraph on the same page as the block after it. */
	keepNext?: boolean;
	/** `w:keepLines`: never split this paragraph's own lines across pages. */
	keepLines?: boolean;
	/** `w:widowControl`; Word's document default is on. Undefined means on. */
	widowControl?: boolean;
	/** Floating pictures anchored in this paragraph. */
	floats?: LayoutFloat[];
	/** Custom tab stops (`w:tabs`). */
	tabStops?: LayoutTabStop[];
	/** Footnotes referenced here. Runs identify their marker; unanchored notes use the first line. */
	footnotes?: LayoutNote[];
	/** Border lines around the paragraph, with each side's gap to the text. */
	borders?: LayoutParagraphBorders;
	/** `#rrggbb` shading behind the paragraph. */
	shading?: string;
	/** `w:kinsoku w:val="0"`: no line-start or line-end prohibition (on when absent). */
	kinsoku?: false;
	/** `w:autoSpaceDE w:val="0"` / `w:autoSpaceDN w:val="0"`: no space between East Asian text and Western letters / digits. */
	autoSpaceDE?: false;
	autoSpaceDN?: false;
	/** `w:overflowPunct w:val="0"`: no line-end punctuation hangs past the margin. */
	overflowPunct?: false;
	/** The document's kinsoku settings (settings.xml), copied onto each paragraph that uses them. */
	kinsokuSettings?: import('../../text/kinsoku').KinsokuSettings;
}

/** A resolved cell border, ready to draw. */
export interface LayoutBorder {
	widthPx: number;
	style: 'solid' | 'double' | 'dotted' | 'dashed';
	color: string;
}
export interface LayoutCellBorders {
	top?: LayoutBorder;
	right?: LayoutBorder;
	bottom?: LayoutBorder;
	left?: LayoutBorder;
}
export interface LayoutCellPadding {
	top: number;
	right: number;
	bottom: number;
	left: number;
}

export interface LayoutTableCell {
	paragraphs: LayoutParagraph[];
	/** Source merge provenance, retained for bounded paragraph-keeping admission. */
	gridSpan?: number;
	verticalMerge?: 'restart' | 'continue';
	/** Rows a vertical merge starting at this cell covers (2 or more), for vertical alignment. */
	rowSpan?: number;
	/** Source child order including read-only nested table stories. */
	blocks?: LayoutBlock[];
	/** Column width in CSS pixels; undefined cells share the row's remaining width evenly. */
	widthPx?: number;
	/** Left edge from the table's left, from the table grid (merged cells span grid columns). */
	xPx?: number;
	/** AutoFit input: the preferred width (`w:tcW`) in pixels, or as a fraction of the table width. */
	preferredWidthPx?: number;
	preferredFraction?: number;
	/** Cell margins (`w:tcMar`, else the table default); content is inset by them. */
	padding?: LayoutCellPadding;
	borders?: LayoutCellBorders;
	/** `#rrggbb` shading fill. */
	shading?: string;
	verticalAlign?: 'top' | 'center' | 'bottom';
}

export interface LayoutTableRow {
	cells: LayoutTableCell[];
	/** Source height provenance, including zero heights not retained by the existing box adapter. */
	heightSpecified?: boolean;
	/** `w:trPr/w:cantSplit`: this row must stay on one page. */
	cantSplit?: boolean;
	/** `w:trHeight`: a minimum (`atLeast`) or fixed (`exact`) row height. */
	heightPx?: number;
	heightRule?: 'atLeast' | 'exact';
	/** `w:trPr/w:tblHeader`: repeat this row at the top of every page the table continues onto. */
	isHeader?: boolean;
	/** Count the text of vertically merged cells in this row's height (the merge is not placed whole). */
	mergedTextInRow?: boolean;
}

export interface LayoutTable {
	kind: 'table';
	id: string;
	rows: LayoutTableRow[];
	/** Total grid width, for aligning the table in its column. */
	widthPx?: number;
	/** `w:tblInd`: offset from the column's left edge. */
	indentPx?: number;
	/** `w:jc` on the table. */
	alignment?: 'left' | 'center' | 'right';
	/**
	 * Set for an AutoFit table (`w:tblLayout` not `fixed`): its columns are sized from their
	 * content when it is placed (table-autofit.ts), replacing the grid widths above.
	 */
	autofit?: LayoutTableAutofit;
	/** `w:tblpPr`: the table is placed out of the text flow and text wraps around it. */
	floating?: LayoutTableFloating;
}

/** A floating table's frame, in the terms floating pictures use (floats.ts). */
export interface LayoutTableFloating {
	relativeFromH: 'page' | 'margin' | 'column';
	alignH?: 'left' | 'center' | 'right' | 'inside' | 'outside';
	offsetXPx: number;
	/** `paragraph` is the paragraph that follows the table, where the table sits in the flow. */
	relativeFromV: 'page' | 'margin' | 'paragraph';
	alignV?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
	offsetYPx: number;
	/** Distance kept between the table and the text wrapped around it. */
	distance: { topPx: number; rightPx: number; bottomPx: number; leftPx: number };
}

export interface LayoutTableAutofit {
	/** Grid columns. */
	columns: number;
	/** Preferred table width (`w:tblW`): pixels, or a fraction of the text column. */
	widthPx?: number;
	widthFraction?: number;
	/** Compatibility modes before 15 let the cells' outer margins overhang the text column. */
	legacyMargins?: boolean;
}

export type LayoutBlock = LayoutParagraph | LayoutTable;

export interface LayoutPageGeometry {
	widthPx: number;
	heightPx: number;
	marginTopPx: number;
	marginRightPx: number;
	marginBottomPx: number;
	marginLeftPx: number;
}

export interface LayoutColumns {
	count: number;
	gapPx: number;
	separator?: boolean;
	/** Explicit column widths and the gap after each column (`w:col`). */
	widths?: { widthPx: number; gapPx: number }[];
}

export interface LayoutSection {
	/** Native mode15 column area, independent of the reference body column. */
	footnoteColumns?: LayoutColumns;
	/** Native Word mode15 repeats the normal story on continued pages. */
	modernFootnoteSeparator?: boolean;
	/** `doNotUseHTMLParagraphAutoSpacing`: add the space after and before instead of the larger. */
	addParagraphSpacing?: boolean;
	/** Compatibility mode before 15: `pageBreakBefore` keeps the space before at the page top. */
	keepSpaceBeforePageBreak?: boolean;
	/** Space after (px) of the previous section's last paragraph, which collapses with this one's first. */
	leadingSpacingAfterPx?: number;
	/** Measured Word compatibility modes 12 and 14 attach notes to body columns. */
	legacyFootnoteColumns?: boolean;
	footnoteSeparators?: { separator?: LayoutParagraph[]; continuationSeparator?: LayoutParagraph[] };
	page: LayoutPageGeometry;
	columns?: LayoutColumns;
	blocks: LayoutBlock[];
	/** How this section begins relative to the previous one; the first section is always a fresh page. */
	break?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
	/** `w:vAlign`: where content sits vertically on each page of the section. */
	verticalAlign?: 'top' | 'center' | 'both' | 'bottom';
}

export interface LayoutDocumentInput {
	sections: LayoutSection[];
}
