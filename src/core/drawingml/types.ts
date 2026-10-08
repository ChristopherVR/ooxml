// Format-neutral DrawingML model (ECMA-376 Part 1, 20.1): colours, fills, lines, text bodies and
// geometry as written in a part, before any theme is applied. No slide, document, sheet or diagram
// types appear here; `diagram`, `chart`, `docx`, `xlsx` and (later) `pptx` adapt to it.

/** Reads one attribute by local name; `undefined` when absent. Lets one parser serve a DOM element and an object tree. */
export type AttributeReader = (name: string) => string | undefined;

/** A theme-independent DrawingML colour as written in a part. */
export interface DrawingColor {
	kind: 'srgb' | 'scheme' | 'system' | 'preset' | 'scrgb' | 'hsl';
	/** Hex for `srgb`, the scheme/preset/system name, or the serialized channels for `scrgb`/`hsl`. */
	value: string;
	/** Colour transforms in document order (`alpha`, `lumMod`, `shade`...), raw attribute values. */
	transforms: { name: string; value: string }[];
	/** `lastClr` of a system colour: the value the producer saw. */
	fallback?: string;
}

/** A DrawingML fill (`a:noFill`, `a:solidFill`, `a:gradFill`, `a:pattFill`, `a:blipFill`). */
export type DrawingFill =
	| { kind: 'none' }
	| { kind: 'solid'; color: DrawingColor }
	| {
			kind: 'gradient';
			/** Original gradient XML, retaining flags and unsupported properties for editing. */
			sourceXml?: string;
			stops: { position: number; color: DrawingColor }[];
			/** Linear angle in degrees, when `a:lin` is present. */
			angle?: number;
			/** Whether the linear vector scales with the shape's bounding box. */
			scaled?: boolean;
			path?: string;
			/** DrawingML path-gradient focus rectangle, as fractions of the shape box. */
			fillToRect?: { l: number; t: number; r: number; b: number };
			/** Gradient tile insets; corner directions can extend beyond the shape box. */
			tileRect?: { l: number; t: number; r: number; b: number };
	  }
	| { kind: 'pattern'; preset: string; foreground?: DrawingColor; background?: DrawingColor }
	| { kind: 'picture'; relId?: string }
	| { kind: 'unsupported'; element: string };

/** A DrawingML outline (`a:ln`). */
export interface DrawingLine {
	/** Width in EMU. */
	widthEmu?: number;
	fill?: DrawingFill;
	dash?: string;
	cap?: string;
}

/** One run of a DrawingML text body. */
export interface DrawingTextRun {
	text: string;
	/** Font size in points. */
	sizePt?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	color?: DrawingColor;
	typeface?: string;
}

/** DrawingML spacing, retaining absolute points or a fractional percentage. */
export interface DrawingTextSpacing {
	unit: 'points' | 'percent';
	value: number;
}

export interface DrawingTextParagraph {
	/** Paragraph default run properties, inherited by individual runs. */
	defaultProperties?: Omit<DrawingTextRun, 'text'>;
	/** `l`, `ctr`, `r`, `just`... */
	align?: string;
	lineSpacing?: DrawingTextSpacing;
	spaceBefore?: DrawingTextSpacing;
	spaceAfter?: DrawingTextSpacing;
	runs: DrawingTextRun[];
}

/** A DrawingML text body (`a:txBody`, `dsp:txBody`, `c:txPr`...). */
export interface DrawingTextBody {
	/** `t`, `ctr`, `b`... */
	anchor?: string;
	/** Insets in EMU: left, top, right, bottom. */
	insetsEmu?: { left?: number; top?: number; right?: number; bottom?: number };
	paragraphs: DrawingTextParagraph[];
	/** Paragraph texts joined by a newline. */
	text: string;
}

/** Position and size in EMU. */
export interface DrawingFrame {
	x: number;
	y: number;
	width: number;
	height: number;
}

export type DrawingPathCommand =
	| { op: 'M' | 'L'; x: number; y: number }
	| { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { op: 'Q'; x1: number; y1: number; x: number; y: number }
	| { op: 'A'; wR: number; hR: number; startAngle: number; swingAngle: number }
	| { op: 'Z' };

/** One `a:path` of a custom geometry, in the path's own coordinate space. */
export interface DrawingPath {
	width?: number;
	height?: number;
	fill?: string;
	stroke?: boolean;
	commands: DrawingPathCommand[];
}
