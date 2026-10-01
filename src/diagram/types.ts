// Format-neutral SmartArt (DiagramML) model. No PowerPoint, Word, slide or document types appear
// here: both `pptx` and `docx` adapt to it. Extracted from `pptx/core/types/smart-art*.ts`.

/** Resolved SmartArt layout family. */
export type DiagramLayoutType =
	| 'list'
	| 'process'
	| 'cycle'
	| 'hierarchy'
	| 'relationship'
	| 'matrix'
	| 'pyramid'
	| 'funnel'
	| 'gear'
	| 'target'
	| 'timeline'
	| 'venn'
	| 'chevron'
	| 'bending'
	| 'unknown';

/** SmartArt colour scheme presets. */
export type DiagramColorScheme =
	| 'colorful1'
	| 'colorful2'
	| 'colorful3'
	| 'monochromatic1'
	| 'monochromatic2';

/** SmartArt visual style intensity. */
export type DiagramStyleIntensity = 'flat' | 'moderate' | 'intense';

/** A connection (`dgm:cxn`) between two data-model points. */
export interface DiagramConnection {
	/** Stable CT_Cxn model identifier. */
	modelId?: string | null;
	sourceId: string;
	destId: string;
	/** `parOf` (the schema default when omitted), `presOf`, `presParOf`, `unknownRelationship`. */
	type?: string;
	srcOrd?: number;
	destOrd?: number;
	/** Model ID of the parent transition point associated with this edge. */
	parentTransitionId?: string | null;
	/** Model ID of the sibling transition point associated with this edge. */
	siblingTransitionId?: string | null;
	/** Layout presentation identifier used by presentation connections. */
	presentationId?: string | null;
	/** Text typed onto the linked `parTrans`/`sibTrans` transition point, when it carries any. */
	label?: string;
}

/** Manual layout overrides (`dgm:prSet/@cust*`) a user applied to a presentation point. */
export interface DiagramNodeCustomLayout {
	/** `custAng`: additional rotation in degrees. */
	angle?: number;
	scaleX?: number;
	scaleY?: number;
	sizeX?: number;
	sizeY?: number;
	flipHorizontal?: boolean;
	flipVertical?: boolean;
	linearFactorX?: number;
	linearFactorY?: number;
	linearFactorNeighborX?: number;
	linearFactorNeighborY?: number;
	radialScaleRadius?: number;
	radialScaleIncrement?: number;
	hasCustomTransform?: boolean;
}

/** The four required relationship ids on `dgm:relIds` (`r:dm`, `r:lo`, `r:qs`, `r:cs`). */
export interface DiagramRelationshipIds {
	dataRelId?: string;
	layoutRelId?: string;
	styleRelId?: string;
	colorsRelId?: string;
}

/** Reads one attribute by local name; `undefined` when absent. Lets one parser serve a DOM element and an object tree. */
export type AttributeReader = (name: string) => string | undefined;

/** A recoverable problem found while reading a diagram part. */
export interface DiagramIssue {
	code: string;
	message: string;
	part?: string;
}

/** A `dgm:pt` of the data model. */
export interface DiagramPoint {
	modelId: string;
	/** `node` (the schema default), `doc`, `asst`, `pres`, `parTrans`, `sibTrans`. */
	type: string;
	/** Paragraphs of the point's `dgm:t` text body, each the concatenation of its runs. */
	paragraphs: string[];
	/** The paragraphs joined by a newline. */
	text: string;
	connectionId?: string;
	/** `dgm:prSet` identifiers that tie a presentation point to its data point and layout node. */
	presAssocId?: string;
	presName?: string;
	presStyleLabel?: string;
	presStyleIndex?: number;
	presStyleCount?: number;
	/** Layout, quick style and colour definitions chosen on the document point's `prSet`. */
	layoutTypeId?: string;
	layoutCategoryId?: string;
	quickStyleTypeId?: string;
	colorsTypeId?: string;
	customLayout?: DiagramNodeCustomLayout;
}

/** A parsed `dgm:dataModel`. */
export interface DiagramDataModel {
	points: DiagramPoint[];
	connections: DiagramConnection[];
	/** Content points (not `doc`, `pres`, `parTrans`, `sibTrans`) in document order. */
	nodes: DiagramPoint[];
	/** `destId -> srcId` for `parOf` connections (first one wins). */
	parentById: Map<string, string>;
	/** `relId` of `dsp:dataModelExt`: the relationship (in the host part) of the cached drawing. */
	drawingRelId?: string;
	issues: DiagramIssue[];
}

/** Header shared by layout, colours and quick-style definition parts. */
export interface DiagramDefinitionHeader {
	uniqueId: string;
	title?: string;
	description?: string;
	categories: { type: string; priority?: number }[];
}

/** A theme-independent DrawingML colour as written in a part. */
export interface DiagramColor {
	kind: 'srgb' | 'scheme' | 'system' | 'preset' | 'scrgb' | 'hsl';
	/** Hex for `srgb`, the scheme/preset/system name, or the serialized channels for `scrgb`/`hsl`. */
	value: string;
	/** Colour transforms in document order (`alpha`, `lumMod`, `shade`...), raw attribute values. */
	transforms: { name: string; value: string }[];
	/** `lastClr` of a system colour: the value the producer saw. */
	fallback?: string;
}

/** A colour choice list (`dgm:fillClrLst` and friends). */
export interface DiagramColorList {
	/** `span`, `cycle` or `repeat`. */
	method?: string;
	colors: DiagramColor[];
}

export interface DiagramColorStyleLabel {
	name: string;
	fill: DiagramColorList;
	line: DiagramColorList;
	effect: DiagramColorList;
	textLine: DiagramColorList;
	textFill: DiagramColorList;
	textEffect: DiagramColorList;
}

/** A parsed `dgm:colorsDef`. */
export interface DiagramColorsDefinition extends DiagramDefinitionHeader {
	labels: DiagramColorStyleLabel[];
}

/** Style-matrix reference (`a:lnRef`, `a:fillRef`, `a:effectRef`, `a:fontRef`). */
export interface DiagramStyleRef {
	/** `idx` of line, fill and effect references. */
	index?: number;
	/** `idx` of a font reference (`major`, `minor`, `none`). */
	fontIndex?: string;
	color?: DiagramColor;
}

export interface DiagramStyleLabel {
	name: string;
	line?: DiagramStyleRef;
	fill?: DiagramStyleRef;
	effect?: DiagramStyleRef;
	font?: DiagramStyleRef;
	/** The label carries a camera and light rig (`dgm:scene3d`). */
	hasScene3d: boolean;
	/** The label carries a non-empty `dgm:sp3d` (bevel, extrusion, material). */
	hasShape3d: boolean;
}

/** A parsed `dgm:styleDef` (quick style). */
export interface DiagramQuickStyleDefinition extends DiagramDefinitionHeader {
	labels: DiagramStyleLabel[];
	/** Any label has a shape 3D definition. */
	has3d: boolean;
}

/** A parsed `dgm:layoutDef`: the header and root algorithm only; the layout tree is not modelled here yet. */
export interface DiagramLayoutSummary extends DiagramDefinitionHeader {
	family?: DiagramLayoutType;
	rootAlgorithm?: string;
	/** Number of `dgm:layoutNode` elements. */
	layoutNodeCount: number;
}

/** A fill of a cached drawing shape. */
export type DiagramFill =
	| { kind: 'none' }
	| { kind: 'solid'; color: DiagramColor }
	| {
			kind: 'gradient';
			stops: { position: number; color: DiagramColor }[];
			/** Linear angle in degrees, when `a:lin` is present. */
			angle?: number;
			path?: string;
	  }
	| { kind: 'pattern'; preset: string; foreground?: DiagramColor; background?: DiagramColor }
	| { kind: 'picture'; relId?: string }
	| { kind: 'unsupported'; element: string };

export interface DiagramLine {
	/** Width in EMU. */
	widthEmu?: number;
	fill?: DiagramFill;
	dash?: string;
	cap?: string;
}

export interface DiagramStyleReference {
	line?: DiagramStyleRef;
	fill?: DiagramStyleRef;
	effect?: DiagramStyleRef;
	font?: DiagramStyleRef;
}

/** One run of a drawing shape's text. */
export interface DiagramTextRun {
	text: string;
	/** Font size in points. */
	sizePt?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	color?: DiagramColor;
	typeface?: string;
}

export interface DiagramTextParagraph {
	/** `l`, `ctr`, `r`, `just`... */
	align?: string;
	runs: DiagramTextRun[];
}

export interface DiagramTextBody {
	/** `t`, `ctr`, `b`... */
	anchor?: string;
	/** Insets in EMU: left, top, right, bottom. */
	insetsEmu?: { left?: number; top?: number; right?: number; bottom?: number };
	paragraphs: DiagramTextParagraph[];
	/** Paragraph texts joined by a newline. */
	text: string;
}

/** Position and size in EMU, relative to the diagram frame. */
export interface DiagramFrame {
	x: number;
	y: number;
	width: number;
	height: number;
}

export type DiagramPathCommand =
	| { op: 'M' | 'L'; x: number; y: number }
	| { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { op: 'Q'; x1: number; y1: number; x: number; y: number }
	| { op: 'A'; wR: number; hR: number; startAngle: number; swingAngle: number }
	| { op: 'Z' };

export interface DiagramPath {
	width?: number;
	height?: number;
	fill?: string;
	stroke?: boolean;
	commands: DiagramPathCommand[];
}

/** A shape of the cached `dsp:drawing`, renderer-neutral. Geometry in EMU, angles in degrees. */
export interface DiagramDrawingShape {
	/** `dsp:sp/@modelId`: the presentation point this shape was laid out from. */
	modelId: string;
	frame: DiagramFrame;
	rotation?: number;
	flipHorizontal?: boolean;
	flipVertical?: boolean;
	/** Preset geometry name (`roundRect`), or `custom` when the shape has `a:custGeom`. */
	geometry: string;
	adjustments?: Record<string, number>;
	/** Custom geometry paths in the path's own coordinate space. */
	paths?: DiagramPath[];
	fill?: DiagramFill;
	line?: DiagramLine;
	style?: DiagramStyleReference;
	text?: DiagramTextBody;
	/** `dsp:txXfrm`: the independent text rectangle. */
	textFrame?: DiagramFrame;
	/** The shape carries `a:scene3d` or `a:sp3d`; not modelled, kept only in the source part. */
	has3d: boolean;
}

/** A parsed `dsp:drawing`: the layout the producing application last computed, cached in the package. */
export interface DiagramDrawing {
	shapes: DiagramDrawingShape[];
	issues: DiagramIssue[];
}
