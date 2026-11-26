// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Inline pictures and hyperlink targets carried by text runs.
import type { DocxChart } from './chart';
import type { DocxDiagram } from './diagram';
import type { StRelFromH, StRelFromV } from './generated/wml-simple-types';
import type { DrawingTransform } from './drawing-arrange';
import type { ShapeSpec } from './shape-model';
import type { DrawingBodyProperties } from '../drawingml/text-model';
import type { Block } from './model';

/** Native payload kept alongside the preview without exposing model or application bytes as a picture. */
export type InlineDrawingPayload =
	| { kind: 'model3d'; partName: string }
	| { kind: 'officePackage'; partName: string; progId: string };

/**
 * A text box's `w:txbxContent`, read by the same block parser as the document body, with the
 * layout settings of its `wps:bodyPr`: anchor, insets, `vert` (other than `horz`), `wrap="none"`
 * and the autofit kind. Display only: the box XML is preserved on save.
 */
export interface TextBoxContent extends Pick<
	DrawingBodyProperties,
	'anchor' | 'insetLeftEmu' | 'insetTopEmu' | 'insetRightEmu' | 'insetBottomEmu' | 'vertical'
> {
	blocks: Block[];
	wordArt?: import('../drawingml/word-art').WordArtStyle;
	/** `wrap="none"`: lines do not wrap, and Word sizes the shape's width to the longest line. */
	wrap?: 'none';
	/** `a:spAutoFit` (`shape`): the shape's height follows its text. Word draws `normAutofit` as none. */
	autofit?: 'shape' | 'normal' | 'none';
	/**
	 * A linked text box chain (`wps:txbx/@id` on the first box, `wps:linkedTxbx` on the others): the
	 * chain id and this box's place in it (0 for the box holding the text). A later box has no
	 * blocks of its own; text that overflows one box continues in the next.
	 */
	link?: { id: number; seq: number };
}

/** An inline drawing (`w:drawing` or legacy `w:pict`) modeled at run granularity. */
export interface InlineImage {
	/** Relationship id in `word/_rels/document.xml.rels` pointing at the media part. */
	relId: string;
	/** Package part name holding the image bytes, e.g. `word/media/image1.png`. */
	partName: string;
	contentType: string;
	/** Shared source-backed crop, placement and ordered picture effects. */
	pictureFill?: Extract<import('../drawingml/types').DrawingFill, { kind: 'picture' }>;
	/** DrawingML effect list (`pic:spPr/a:effectLst`), e.g. picture drop shadows. */
	effects?: import('../drawingml/effect-model').DrawingEffectList;
	/** An SVG original (`asvg:svgBlip`) shown instead of the raster `partName`, which is its PNG fallback. */
	svgPartName?: string;
	/** Display size in CSS pixels (96 dpi), derived from the drawing extent in EMU (may be fractional). */
	widthPx: number;
	heightPx: number;
	/** Picture rotation and flips, from its DrawingML transform. */
	transform?: DrawingTransform;
	/** From `wp:docPr/@descr`. */
	altText?: string;
	/** A preferred `wp:docPr/@id` when promoting a group child to its own drawing. */
	drawingId?: string;
	/** From `wp:docPr/@name`. */
	drawingName?: string;
	/** From `wp:docPr/@title`. */
	title?: string;
	/** Whether the drawing has a `wp:anchor` and participates in page float layout. */
	anchored?: boolean;
	/** An absolute legacy VML picture with source-backed placement. */
	vmlFloat?: true;
	/** Floating wrapping, position and stacking order. */
	placement?: PicturePlacement;
	/** Set for non-picture drawings (chart, SmartArt, shape, unresolved legacy VML): rendered as a labeled placeholder with no editable bytes. */
	unsupported?: string;
	/** For a text box (`wps:txbx`): its paragraphs' text, shown read-only; the box XML is preserved on save. */
	textBoxText?: string[];
	/** The text box is the simple form this editor writes, so its text, size and outline can change. */
	watermark?: WatermarkSpec;
	/** A Word picture watermark (`WordPictureWatermark`); `partName` and `contentType` name its picture. */
	pictureWatermark?: PictureWatermarkSpec;
	textBoxEditable?: boolean;
	/** For an editable text box: whether it draws an outline (default true). */
	textBoxBorder?: boolean;
	/** For a drawing shape (`unsupported: "Shape"`): its preset geometry, fill, outline and text. */
	shape?: ShapeSpec;
	/** For a shape or text box with a `wps:txbx`: its content as paragraphs and tables. */
	textBoxContent?: TextBoxContent;
	/** For SmartArt (`unsupported: "SmartArt"`): its parts, extent and cached drawing; the markup and parts are preserved on save. */
	diagram?: DocxDiagram;
	/** For a chart (`unsupported: "Chart"`): its part and the parsed `c:chartSpace`; the markup and parts are preserved on save. */
	chart?: DocxChart;
	/** A Word signature line (`o:signatureline` in a VML picture): its setup, drawn as the picture. */
	signatureLine?: import('./signature-line').SignatureLineSpec;
	/** An online video (`wp15:webVideoPr`): the embed it plays, drawn as its poster picture. */
	webVideo?: import('./web-video').WebVideoSpec;
	/** A native 3D model or OLE package. The picture fields name its poster or icon. */
	payload?: InlineDrawingPayload;
	/** Ink (`unsupported: "Ink"`): its strokes and InkML part; drawn and edited on the Draw tab. */
	ink?: import('./ink').DocxInk;
}
/** A Word text watermark (`PowerPlusWaterMarkObject`), carried by an inline object run in a header. */
export interface WatermarkSpec {
	text: string;
	fontFamily?: string;
	/** `#rrggbb`. */
	color: string;
	/** Drawn at half opacity. */
	semitransparent: boolean;
	layout: 'diagonal' | 'horizontal';
	/** Shape width in points (from VML `style`), when read from a document. */
	widthPt?: number;
	/** Shape height in points (from VML `style`), when read from a document. */
	heightPt?: number;
}
/**
 * A Word picture watermark: a VML picture centred on the margins behind the text of every page,
 * carried by an inline object run in a header.
 */
export interface PictureWatermarkSpec {
	kind: 'picture';
	partName: string;
	contentType: string;
	/** Drawn size in points. */
	widthPt: number;
	heightPt: number;
	/** The Scale chosen in the dialog: `auto` fits the text area, a number is a percentage of the
	 *  picture's own size. Undefined when read from a file, which does not record it. */
	scale?: 'auto' | number;
	/** Word's Washout: the picture drawn light (`v:imagedata` gain and blacklevel). */
	washout: boolean;
	/** `v:imagedata/@o:title`, the picture's file name without extension. */
	title?: string;
}
/** How a floating picture sits relative to text. */
export interface PicturePlacement {
	/** Unsigned relative stacking order within the same text layer. */
	relativeHeight?: number;
	wrap: 'square' | 'tight' | 'through' | 'topAndBottom' | 'none';
	/** `behindDoc`: with `wrap: 'none'`, the picture is behind rather than in front of text. */
	behindText?: boolean;
	/** `wrapText` of a square, tight or through wrap: which sides of the picture text may use. */
	wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
	/** `wp:positionH/wp:align`, when the picture is aligned rather than offset. */
	align?: 'left' | 'center' | 'right' | 'inside' | 'outside';
	/** `wp:positionH/wp:posOffset` in CSS pixels, relative to `relativeFrom`. */
	offsetXPx?: number;
	relativeFrom?: StRelFromH;
	/** `wp:positionV/wp:align`, when the picture is aligned vertically rather than offset. */
	alignV?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
	/** `wp:positionV/wp:posOffset` in CSS pixels, relative to `relativeFromV`. */
	offsetYPx?: number;
	/** `wp:positionV/@relativeFrom`, e.g. `paragraph`, `page` or `margin`. */
	relativeFromV?: StRelFromV;
	/**
	 * `layoutInCell="0"`: an object anchored in a table cell is positioned on the page instead of in
	 * the cell. Word honours it only before compatibility mode 15; absent means the default, `1`.
	 */
	layoutInCell?: false;
}
/** A `w:hyperlink` target, resolved from its relationship (external) or `w:anchor` (internal). */
export interface HyperlinkInfo {
	href?: string;
	anchor?: string;
	tooltip?: string;
}
