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
	DiagramShapeView,
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
 * A pre-computed shape from `ppt/diagrams/drawing*.xml`: the neutral
 * {@link DiagramShapeView} (geometry in pixels, resolved colours, text style)
 * plus the pptx rich-text segments, 3D and custom-geometry slots.
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
export interface PptxSmartArtDrawingShape extends DiagramShapeView, PptxCustomPathProperties {
	/** Standard rich-text segments projected from the associated SmartArt node. */
	textSegments?: TextSegment[];
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
