// Format-neutral SmartArt (DiagramML) model. No PowerPoint, Word, slide or document types appear
// here: both `pptx` and `docx` adapt to it. Extracted from `pptx/core/types/smart-art*.ts`.

import type {
	DrawingColor,
	DrawingFill,
	DrawingFrame,
	DrawingLine,
	DrawingPath,
	DrawingPathCommand,
	DrawingTextBody,
	DrawingTextParagraph,
	DrawingTextRun,
	DrawingTextSpacing,
} from '../drawingml/types';

export type { AttributeReader } from '../drawingml/types';

// The DrawingML primitives a diagram uses are the neutral `drawingml` types under their
// historical diagram names, kept so the public `ooxml-core/diagram` API does not change.
export type DiagramColor = DrawingColor;
export type DiagramFill = DrawingFill;
export type DiagramLine = DrawingLine;
export type DiagramTextRun = DrawingTextRun;
export type DiagramTextSpacing = DrawingTextSpacing;
export type DiagramTextParagraph = DrawingTextParagraph;
export type DiagramTextBody = DrawingTextBody;
export type DiagramFrame = DrawingFrame;
export type DiagramPathCommand = DrawingPathCommand;
export type DiagramPath = DrawingPath;

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

export interface DiagramStyleReference {
	line?: DiagramStyleRef;
	fill?: DiagramStyleRef;
	effect?: DiagramStyleRef;
	font?: DiagramStyleRef;
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
