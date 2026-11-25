// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// A drawing shape (`wps:wsp` with a preset geometry): what the editor draws, edits and writes. The
// geometry, fill and outline are the format-neutral DrawingML types, so the same readers and
// writers serve every format.
import type { DrawingColor, DrawingEffectList, DrawingFill, DrawingLine } from '../drawingml/index';

/** One child shape inside a WordprocessingML group (`wpg:wgp`), in group-child pixel space. */
export interface GroupShapeChild {
	xPx: number;
	yPx: number;
	widthPx: number;
	heightPx: number;
	shape: ShapeSpec;
}

export interface ShapeSpec {
	/** `a:prstGeom/@prst`, e.g. `rect`, `roundRect`, `rightArrow`, `line`. */
	preset: string;
	/** Source-backed custom geometry; retained verbatim through formatting and resize edits. */
	customGeometryXml?: string;
	/** Child shapes when this drawing is a `wpg:wgp` group. */
	groupChildren?: GroupShapeChild[];
	/** Group child coordinate space (`a:xfrm/a:chOff`, `a:chExt`) in CSS pixels. */
	groupExtent?: { xPx: number; yPx: number; widthPx: number; heightPx: number };
	/** `a:avLst` guide values by name (Word's yellow handles), in the guide's own units. */
	adjustments?: Record<string, number>;
	/** `spPr` fill. Undefined takes the shape style's fill (the theme's first accent colour). */
	fill?: DrawingFill;
	/** `spPr/a:ln`. Undefined takes the shape style's outline (a darker accent). */
	line?: DrawingLine;
	/** `spPr/a:effectLst` (outer shadows and other DrawingML effects). */
	effects?: DrawingEffectList;
	/** `wps:style/a:fontRef` colour: the colour of the shape's text where its runs set none. */
	fontColor?: DrawingColor;
	/** Paragraphs of text inside the shape (`wps:txbx`). */
	text?: string[];
	/** Uniform direct size of a grouped shape's runs, in points. */
	groupTextSizePt?: number;
	/**
	 * `RRGGBB`: the one `w:color` every run of the shape's plain text sets, written back on its
	 * runs when the text is rewritten. Absent when the runs set none.
	 */
	textColor?: string;
	/** The shape's spPr holds only what this model writes, so fill, outline, size, geometry and wrap can change. */
	editable?: boolean;
	/** The shape's text is plain paragraphs of plain runs, so it can be edited. */
	textEditable?: boolean;
	/** Counterclockwise-free `a:xfrm/@rot` in degrees, read and kept; drawn rotated. */
	rotation?: number;
	flipH?: boolean;
	flipV?: boolean;
}

/** Office's default shape colours: the theme's first accent, and the same darkened by half. */
export const DEFAULT_SHAPE_FILL = '#4472C4';
export const DEFAULT_SHAPE_LINE = '#2F528F';
/** The default outline is 1 pt (12700 EMU). */
export const DEFAULT_SHAPE_LINE_EMU = 12700;

/** A shape in Word's default style, ready to place. */
export function newShape(preset: string): ShapeSpec {
	return { preset, editable: true, textEditable: true };
}

/** Whether the preset is an open line or connector: no fill, drawn as an outline only. */
export const isLinePreset = (preset: string): boolean =>
	/^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(preset);
