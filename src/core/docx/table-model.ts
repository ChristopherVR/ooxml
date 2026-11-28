import type { StTextDirection } from './generated/wml-simple-types';
import type { StyleMetadata } from './style-metadata';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, Paragraph, ParagraphFormatting } from './model';
import type { RunFormatting } from './run-style-model';
import type { ThemeColorReference, ThemeColorToken } from './theme-model';
import type { StBorder, StVerticalJc } from './generated/wml-simple-types';
import type { EighthPoints, SignedTwips, Twips } from './units';

export interface TableBorderSide {
	style?: StBorder;
	/** Border width in eighths of a point (`w:sz`). */
	sizeEighthPoints?: EighthPoints;
	color?: string;
	themeColor?: ThemeColorToken;
	/** `w:themeTint` / `w:themeShade` of the theme colour, as 0..1 fractions of the byte. */
	themeTint?: number;
	themeShade?: number;
	/** Paragraph borders only: gap between the border and the text (`w:space`), in points. */
	spacePoints?: number;
	/** Paragraph borders only: the border casts a shadow to the right and below (`w:shadow`). */
	shadow?: boolean;
	/** Page borders only: a Word art border (`w:art`); drawn as a plain line. */
	art?: string;
}
/** Paragraph borders (`w:pBdr`); `between` separates paragraphs that share the same borders. */
export interface ParagraphBorders {
	top?: TableBorderSide;
	bottom?: TableBorderSide;
	left?: TableBorderSide;
	right?: TableBorderSide;
	between?: TableBorderSide;
}
export interface TableBorders {
	top?: TableBorderSide;
	bottom?: TableBorderSide;
	left?: TableBorderSide;
	right?: TableBorderSide;
	insideH?: TableBorderSide;
	insideV?: TableBorderSide;
}
export interface TableCellMargins {
	top?: SignedTwips;
	bottom?: SignedTwips;
	left?: SignedTwips;
	right?: SignedTwips;
}
/** `w:tblPr/w:tblpPr`: a floating table's position and the distance kept from wrapped text. */
export interface TableFloatingPosition {
	/** `w:horzAnchor`: `text` is the text column. */
	horizontalAnchor?: 'margin' | 'page' | 'text';
	/** `w:vertAnchor`: `text` is the paragraph that follows the table. */
	verticalAnchor?: 'margin' | 'page' | 'text';
	/** `w:tblpX`/`w:tblpY`: offsets from the anchor; `w:tblpXSpec`/`w:tblpYSpec` win over them. */
	xTwips?: SignedTwips;
	yTwips?: SignedTwips;
	xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
	yAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
	/** `w:leftFromText`, `w:rightFromText`, `w:topFromText`, `w:bottomFromText`. */
	leftFromTextTwips?: Twips;
	rightFromTextTwips?: Twips;
	topFromTextTwips?: Twips;
	bottomFromTextTwips?: Twips;
}
/** Row-level properties (`w:trPr`): height, keep-together and repeat-as-header. */
export interface TableRowProperties {
	/** The `w:sdt` around this row: saving patches what changed, and wraps or unwraps the row. */
	contentControl?: import('./content-control-model').ContentControl;
	/** The row-level controls around `contentControl` (a repeating section around its item). */
	outerContentControls?: import('./content-control-model').ContentControl[];
	heightTwips?: Twips;
	/** `exact` fixes the height; `atLeast` (Word's `atLeast` and `auto`) is a minimum. */
	heightRule?: 'atLeast' | 'exact';
	/** `w:cantSplit`: the row is never split across pages. */
	cantSplit?: boolean;
	/** `w:tblHeader`: the row repeats at the top of each page the table continues onto. */
	header?: boolean;
	/** `w:trPr/w:ins` or `w:trPr/w:del`: the whole row is a tracked insertion or deletion. */
	revision?: import('./model-run').Revision;
	/** `w:trPr/w:trPrChange`: a tracked row-property formatting change. */
	formatRevision?: import('./model-run').Revision;
}
/** A read-only preview of a nested table's text, rendered but not independently editable. */
export interface NestedTablePreview {
	rows: { text: string }[][];
}
export interface TableCell {
	paragraphs: Paragraph[];
	/** The `w:sdt` around this cell: saving patches what changed, and wraps or unwraps the cell. */
	contentControl?: import('./content-control-model').ContentControl;
	/** `w:gridSpan`; number of grid columns this cell occupies (horizontal merge). */
	gridSpan?: number;
	/** `w:vMerge`; `restart` begins a vertical merge, `continue` extends the cell above. */
	verticalMerge?: 'restart' | 'continue';
	widthTwips?: Twips;
	/** A percentage `w:tcW` (`type="pct"`) as a fraction of the table width; read-only layout input. */
	widthPercent?: number;
	verticalAlign?: StVerticalJc;
	textDirection?: StTextDirection;
	/** Layout input from `w:hideMark`: ignore the final cell mark when sizing the row. */
	hideMark?: boolean;
	shadingFill?: string;
	percentageShading?: import('./percentage-shading').PercentageShading;
	shadingThemeFill?: ThemeColorReference;
	borders?: TableBorders;
	margins?: TableCellMargins;
	/** Tables nested directly in this cell's XML, preserved but read-only in the editor. */
	nestedTables?: NestedTablePreview[];
	/** Parsed OOXML child order for cells containing nested tables. Read-only source content. */
	sourceBlocks?: Block[];
	/** `w:tcPr/w:tcPrChange`: a tracked cell-property formatting change. */
	formatRevision?: import('./model-run').Revision;
}
/** Which `tblStyle` conditional formatting regions apply, from `w:tblLook`. */
export interface TableLook {
	firstRow?: boolean;
	lastRow?: boolean;
	firstColumn?: boolean;
	lastColumn?: boolean;
	noHBand?: boolean;
	noVBand?: boolean;
}
export type TableConditionalRegion =
	| 'wholeTable'
	| 'firstRow'
	| 'lastRow'
	| 'firstCol'
	| 'lastCol'
	| 'band1Horz'
	| 'band2Horz'
	| 'band1Vert'
	| 'band2Vert';
/** A conditional-formatting region (`w:tblStylePr`) within a table style. */
export interface TableStyleConditionalFormatting {
	paragraph?: ParagraphFormatting;
	borders?: TableBorders;
	shadingFill?: string;
	percentageShading?: import('./percentage-shading').PercentageShading;
	shadingThemeFill?: ThemeColorReference;
	run?: RunFormatting;
}
export interface TableStyleDefinition extends StyleMetadata {
	paragraphFormatting?: import('./model-paragraph').ParagraphFormatting;
	runFormatting?: RunFormatting;
	shadingThemeFill?: ThemeColorReference;
	borders?: TableBorders;
	shadingFill?: string;
	percentageShading?: import('./percentage-shading').PercentageShading;
	/** `w:tblPr/w:tblCellMar`: default cell margins of tables with this style. */
	cellMargins?: TableCellMargins;
	conditional: Partial<Record<TableConditionalRegion, TableStyleConditionalFormatting>>;
}
/**
 * Parsed `styles.xml` table styles with conditional formatting. Like the
 * paragraph catalog, it is a value: an edit replaces the `styles` record.
 */
export interface TableStyleCatalog {
	styles: Record<string, TableStyleDefinition>;
	warnings: string[];
}
