/**
 * SmartArt types: layout categories, layout presets, colour schemes,
 * data-model nodes/connections, drawing shapes, chrome, and the composite
 * `PptxSmartArtData`.
 *
 * @module pptx-types/smart-art
 */

import type {
	DiagramColorScheme,
	DiagramConnection,
	DiagramData,
	DiagramLayoutType,
	DiagramStyleIntensity,
} from '../../../diagram/index';
import type { PptxCustomPathProperties } from './geometry';
import type { PptxSmartArtChrome } from './smart-art-chrome';
import type {
	PptxSmartArtLayoutDefinition,
	PptxSmartArtNode,
	PptxSmartArtQuickStyle,
} from './smart-art-model';
import type { TextSegment } from './text';
import type { Pptx3DScene, Pptx3DShape, Text3DStyle } from './three-d';

// The neutral SmartArt model (nodes, layout definition, constraints, colour
// transform, quick style, layout variables) lives in `diagram/model`;
// `smart-art-model.ts` binds it to the pptx raw-XML, text-style and 3D types
// under the historical `PptxSmartArt*` names.
export type * from './smart-art-model';
export type * from './smart-art-definition-header';

// ==========================================================================
// SmartArt types
// ==========================================================================

/**
 * Resolved SmartArt layout category.
 *
 * @example
 * ```ts
 * const cat: SmartArtLayoutType = "hierarchy";
 * // => "hierarchy" — one of: "list" | "process" | "cycle" | "hierarchy" | "relationship" | …
 * ```
 */
export type SmartArtLayoutType = DiagramLayoutType;

/**
 * SmartArt colour scheme presets.
 *
 * @example
 * ```ts
 * const scheme: SmartArtColorScheme = "colorful1";
 * // => "colorful1" — one of: "colorful1" | "colorful2" | "colorful3" | "monochromatic1" | "monochromatic2"
 * ```
 */
export type SmartArtColorScheme = DiagramColorScheme;

/**
 * SmartArt visual style intensity.
 *
 * @example
 * ```ts
 * const style: SmartArtStyle = "moderate";
 * // => "moderate" — one of: "flat" | "moderate" | "intense"
 * ```
 */
export type SmartArtStyle = DiagramStyleIntensity;

/**
 * A connection between two SmartArt data-model nodes.
 *
 * @example
 * ```ts
 * const conn: PptxSmartArtConnection = {
 *   sourceId: "1",
 *   destId: "2",
 *   type: "parOf",
 * };
 * // => satisfies PptxSmartArtConnection
 * ```
 */
export type PptxSmartArtConnection = DiagramConnection;

/**
 * A pre-computed shape from `ppt/diagrams/drawing*.xml`.
 *
 * @example
 * ```ts
 * const shape: PptxSmartArtDrawingShape = {
 *   id: "s1",
 *   shapeType: "roundRect",
 *   x: 100, y: 50, width: 200, height: 80,
 *   fillColor: "#4F81BD",
 *   text: "CEO",
 * };
 * // => satisfies PptxSmartArtDrawingShape
 * ```
 */
export interface PptxSmartArtDrawingShape extends PptxCustomPathProperties {
	/** Shape ID within the drawing. */
	id: string;
	/** Preset geometry type (e.g. "roundRect", "ellipse"). */
	shapeType?: string;
	/** Position and size in EMU-based pixels. */
	x: number;
	y: number;
	width: number;
	height: number;
	/** Rotation in degrees. */
	rotation?: number;
	/** Horizontal/vertical mirrors declared on the cached drawing transform. */
	flipHorizontal?: boolean;
	flipVertical?: boolean;
	/** Skew along the X axis in degrees. */
	skewX?: number;
	/** Skew along the Y axis in degrees. */
	skewY?: number;
	/** Preset-geometry adjustment values from `a:prstGeom/a:avLst`. */
	shapeAdjustments?: Record<string, number>;
	/**
	 * The cached shape declares `a:noFill`. Renderers must leave it unpainted
	 * rather than substituting a palette colour, because these shapes usually sit
	 * on top of a painted shape whose fill has to stay visible.
	 */
	fillNone?: boolean;
	/** Solid fill colour (hex). */
	fillColor?: string;
	/**
	 * Opacity (0..1) of the solid fill, from the colour's `a:alpha` (Basic Venn
	 * paints its circles at 50% so overlaps blend). Absent when opaque.
	 */
	fillOpacity?: number;
	/**
	 * Gradient fill stops when the cached shape uses `a:gradFill`. Positions are
	 * 0..100 (percent). Renderers emit an SVG/CSS gradient instead of a flat box.
	 */
	fillGradientStops?: Array<{ color: string; position: number; opacity?: number }>;
	/** Gradient geometry type (`linear` for `a:lin`, `radial` for `a:path`). */
	fillGradientType?: 'linear' | 'radial';
	/** Linear gradient angle in degrees (0..360). */
	fillGradientAngle?: number;
	/** Pattern fill preset name from `a:pattFill/@prst` (e.g. "pct50", "cross"). */
	fillPatternPreset?: string;
	/** Pattern fill foreground colour (hex) from `a:pattFill/a:fgClr`. */
	fillPatternForegroundColor?: string;
	/** Pattern fill background colour (hex) from `a:pattFill/a:bgClr`. */
	fillPatternBackgroundColor?: string;
	/**
	 * Relationship id of a picture (blip) fill's embedded image, from
	 * `a:blipFill/a:blip/@r:embed`. The image bytes are resolved separately; see
	 * {@link fillImageUrl}.
	 */
	fillBlipEmbedId?: string;
	/**
	 * Resolved data-URI/URL for a picture (blip) fill, when the embedded image
	 * part could be resolved. Absent when only {@link fillBlipEmbedId} is known.
	 */
	fillImageUrl?: string;
	/** Whether the cached shape carries an outer-shadow effect (`a:effectLst`). */
	hasShadow?: boolean;
	/** Resolved outer-shadow colour (hex), when present. */
	shadowColor?: string;
	/** Stroke colour (hex). */
	strokeColor?: string;
	/** Stroke width in points. */
	strokeWidth?: number;
	/** Text content of the shape. */
	text?: string;
	/** Standard rich-text segments projected from the associated SmartArt node. */
	textSegments?: TextSegment[];
	/** Font size in CSS pixels. */
	fontSize?: number;
	/** Font colour (hex). */
	fontColor?: string;
	/** Authored font family from the first styled run. */
	fontFamily?: string;
	/** Authored font weight (400 or 700). */
	fontWeight?: number;
	/** Authored font style. */
	fontStyle?: 'normal' | 'italic';
	/** Absolute line height in CSS pixels from `a:spcPts`. */
	lineHeight?: number;
	/** Relative line height from `a:spcPct`. */
	lineHeightRatio?: number;
	/** Absolute spacing after a paragraph in CSS pixels. */
	lineSpacingAfter?: number;
	/** Relative spacing after a paragraph. */
	lineSpacingAfterRatio?: number;
	/** Text-body insets in CSS pixels. */
	textInsetLeft?: number;
	textInsetTop?: number;
	textInsetRight?: number;
	textInsetBottom?: number;
	/** DiagramML text vertical anchor (`t`, `ctr`, or `b`). */
	textVerticalAnchor?: string;
	/** Independent DiagramML text-frame geometry from `dsp:txXfrm`. */
	textFrameX?: number;
	textFrameY?: number;
	textFrameWidth?: number;
	textFrameHeight?: number;
	/**
	 * 3D scene (camera/light rig/backdrop) from `dsp:spPr/a:scene3d`, when this
	 * cached shape carries its own per-shape camera. Bevel quick styles
	 * (Polished, Inset, Cartoon, Powder) cache one per shape (always
	 * `orthographicFront`, so the 2D layout stays undistorted). Scene quick
	 * styles (Brick, Flat, Metallic, Sunset, Bird's Eye) instead put ONE camera
	 * on the whole diagram (see `PptxSmartArtQuickStyle.scene3d`) and leave this
	 * undefined on every shape.
	 */
	scene3d?: Pptx3DScene;
	/**
	 * 3D extrusion/bevel/contour/material from `dsp:spPr/a:sp3d`. PowerPoint
	 * caches this fully resolved (from the quick style's per-label `dgm:sp3d`)
	 * on every shape for every non-flat quick style, bevel or scene alike.
	 */
	shape3d?: Pptx3DShape;
	/**
	 * Text-body 3D extrusion/bevel from `dsp:txBody/a:bodyPr/a:sp3d`. Rare:
	 * present when a quick style extrudes the label text itself off the shape
	 * face (e.g. Bird's Eye Scene: `extrusionH="28000"`), rather than only the
	 * shape body.
	 */
	text3d?: Text3DStyle;
	/**
	 * The presentation style label (`presStyleLbl`: `node1`, `revTx`,
	 * `sibTrans2D1`, ...) of the layout node a REGENERATED shape was laid out
	 * from, when the layout engine reports it. Picks the quick-style label
	 * whose 3D the shape takes; `undefined` on cached (parsed) shapes.
	 */
	styleLabel?: string;
}

// Chrome types (PptxSmartArtChrome, PptxSmartArtRawBackgroundFill) live in
// `smart-art-chrome.ts` to keep this file within the per-file line budget.
export type { PptxSmartArtChrome, PptxSmartArtRawBackgroundFill } from './smart-art-chrome';

/**
 * Complete parsed SmartArt data for a {@link SmartArtPptxElement}: the neutral
 * {@link DiagramData} (bound to the pptx node, quick-style and layout-definition
 * types) plus the package parts, cached drawing and save hints.
 *
 * @example
 * ```ts
 * const data: PptxSmartArtData = {
 *   resolvedLayoutType: "hierarchy",
 *   layout: "hierarchy",
 *   colorScheme: "colorful1",
 *   style: "moderate",
 *   nodes: [
 *     { id: "1", text: "CEO", children: [
 *       { id: "2", text: "VP Marketing", parentId: "1" },
 *     ]},
 *   ],
 * };
 * // => satisfies PptxSmartArtData
 * ```
 */
export interface PptxSmartArtData extends DiagramData {
	nodes: PptxSmartArtNode[];
	/** Pre-computed shapes from `ppt/diagrams/drawing*.xml`. */
	drawingShapes?: PptxSmartArtDrawingShape[];
	/** Background and outline chrome from `dgm:bg` / `dgm:whole`. */
	chrome?: PptxSmartArtChrome;
	/** Quick style from `ppt/diagrams/quickStyles*.xml`. */
	quickStyle?: PptxSmartArtQuickStyle;
	/** Editable metadata from the related DiagramML layout definition. */
	layoutDefinition?: PptxSmartArtLayoutDefinition;
	/** Relationship ID for the diagram data part (for round-trip save). */
	dataRelId?: string;
	/** Relationship ID for the diagram layout part. */
	layoutRelId?: string;
	/** Relationship ID for the drawing part. */
	drawingRelId?: string;
	/** Relationship ID for the colours part. */
	colorsRelId?: string;
	/** Relationship ID for the quick-styles part. */
	styleRelId?: string;
	/**
	 * `uniqueId` of the built-in PowerPoint layout whose definition was applied by a layout swap
	 * (`ooxml-core/pptx/smartart-layouts`). Save writes that definition as the layout part.
	 */
	builtinLayoutId?: string;
	/** Internal save hint: the layout definition changed in the editor. */
	layoutDirty?: boolean;
	/** Internal save hint: typed layout-definition metadata changed. */
	layoutDefinitionDirty?: boolean;
	/** Internal save hint: quick-style definition metadata changed. */
	quickStyleDirty?: boolean;
	/** Internal save hint: color-transform definition metadata changed. */
	colorTransformDirty?: boolean;
	/** Internal save hint: cached drawing geometry or text changed in the editor. */
	drawingDirty?: boolean;
}
