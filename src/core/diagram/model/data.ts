/**
 * The parsed SmartArt model the layout engine reads: data-model nodes and
 * connections, layout variables, the typed layout definition, colour transform
 * and quick style. Package concerns (relationship ids, cached drawing shapes,
 * background chrome, save hints) are added by the format area that loads it
 * (pptx: `PptxSmartArtData extends DiagramData`). Moved from
 * `pptx/core/types/smart-art.ts`.
 */

import type {
	DiagramColorScheme,
	DiagramConnection,
	DiagramLayoutType,
	DiagramStyleIntensity,
} from '../types';
import type { DiagramLayoutDefinition } from './layout-definition';
import type { DiagramNode } from './node';
import type { DiagramColorTransform, DiagramQuickStyle } from './style-definition';

/**
 * Named SmartArt layout presets for creation (subset of the built-in layouts).
 *
 * @example
 * ```ts
 * const layout: DiagramLayoutPreset = "hierarchy";
 * // => "hierarchy" - one of: "basicBlockList" | "alternatingHexagons" | "hierarchy" | ...
 * ```
 */
export type DiagramLayoutPreset =
	| 'basicBlockList'
	| 'alternatingHexagons'
	| 'basicChevronProcess'
	| 'basicCycle'
	| 'basicPie'
	| 'basicRadial'
	| 'basicVenn'
	| 'continuousBlockProcess'
	| 'convergingRadial'
	| 'hierarchy'
	| 'horizontalBulletList'
	| 'linearVenn'
	| 'segmentedProcess'
	| 'stackedList'
	| 'tableList'
	| 'trapezoidList'
	| 'upwardArrow'
	| 'basicFunnel'
	| 'basicTarget'
	| 'interlockingGears'
	| 'basicTimeline'
	| 'basicMatrix'
	| 'basicPyramid'
	| 'invertedPyramid'
	| 'bendingProcess'
	| 'stepDownProcess'
	| 'alternatingFlow'
	| 'descendingProcess'
	| 'pictureAccentList'
	| 'verticalBlockList'
	| 'groupedList'
	| 'pyramidList'
	| 'horizontalPictureList'
	| 'accentProcess'
	| 'verticalChevronList';

/**
 * Presentation layout variables from `dgm:prSet/dgm:presLayoutVars` (data model)
 * or `dgm:varLst` (layout definition defaults).
 *
 * These drive how the DiagramML layout interpreter arranges points: flow
 * direction, hierarchy branch style, org-chart mode, and child count limits.
 * The fallback layout engine can consult them for direction/org-chart hints.
 *
 * @example
 * ```ts
 * const vars: DiagramPresLayoutVars = { direction: "rev", orgChart: true };
 * // => satisfies DiagramPresLayoutVars
 * ```
 */
export interface DiagramPresLayoutVars {
	/** Flow direction (`dgm:dir`): "norm" (default) or "rev" (reversed/RTL). */
	direction?: 'norm' | 'rev';
	/** Hierarchy branch style (`dgm:hierBranch`): std/init/l/r/hang. */
	hierarchyBranch?: 'std' | 'init' | 'l' | 'r' | 'hang';
	/** Org-chart mode enabled (`dgm:orgChart`). */
	orgChart?: boolean;
	/** Maximum children per node (`dgm:chMax`, -1 = unbounded). */
	childMax?: number;
	/** Preferred children per node (`dgm:chPref`, -1 = unbounded). */
	childPreferred?: number;
	/** Whether bullets are enabled (`dgm:bulletEnabled`). */
	bulletEnabled?: boolean;
	/** Animation-by-level setting (`dgm:animLvl`). */
	animationLevel?: string;
	/** Animate-one setting (`dgm:animOne`). */
	animateOne?: string;
	/** Allowed resize handles (`dgm:resizeHandles`). */
	resizeHandles?: string;
}

/**
 * The format-neutral part of a parsed SmartArt graphic: everything the layout
 * engine and interpreters read.
 *
 * @example
 * ```ts
 * const data: DiagramData = {
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
 * // => satisfies DiagramData
 * ```
 */
export interface DiagramData {
	layoutType?: string;
	resolvedLayoutType?: DiagramLayoutType;
	/** Named layout preset (used when creating new SmartArt). */
	layout?: DiagramLayoutPreset;
	/** Colour scheme for the SmartArt graphic. */
	colorScheme?: DiagramColorScheme;
	/** Visual style intensity. */
	style?: DiagramStyleIntensity;
	nodes: DiagramNode[];
	/** Connections between data-model nodes. */
	connections?: DiagramConnection[];
	/** Colour transform from the diagram colours part. */
	colorTransform?: DiagramColorTransform;
	/** Quick style from the diagram quick-style part. */
	quickStyle?: DiagramQuickStyle;
	/** Editable metadata from the related DiagramML layout definition. */
	layoutDefinition?: DiagramLayoutDefinition;
	/**
	 * Presentation layout variables (direction, hierarchy branch, org-chart,
	 * child limits, bullets) from `dgm:presLayoutVars` / layout `dgm:varLst`.
	 * Consulted by the fallback layout engine for direction/org-chart hints.
	 */
	presLayoutVars?: DiagramPresLayoutVars;
	/**
	 * The document theme's minor-Latin font (`a:fontScheme/a:minorFont/a:latin/
	 * @typeface`): what SmartArt text actually renders in when no per-run
	 * `a:latin` override is present (the common case; the font-fit is the one
	 * consumer). Undefined when the theme carries no font scheme at all.
	 */
	themeMinorFont?: string;
}
