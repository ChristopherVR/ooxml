/**
 * SmartArt layout reflow, bound to `PptxSmartArtData`. The reflow algorithms
 * live in `diagram/edit/reflow*` and produce neutral `DiagramShapeView`s,
 * which `PptxSmartArtDrawingShape` extends.
 *
 * @module smartart-editing-reflow
 */

import {
	reflowDiagramLayout,
	resolveDiagramReflowCategory,
	type DiagramReflowBounds,
} from '../../../diagram/edit/index';
import type { PptxSmartArtData, PptxSmartArtDrawingShape, SmartArtLayoutType } from '../types';

/** Position data for a single node produced by the reflow engine. */
export interface ReflowedNodePosition {
	/** Unique identifier of the SmartArt node. */
	nodeId: string;
	/** Horizontal position in pixels. */
	x: number;
	/** Vertical position in pixels. */
	y: number;
	/** Width of the node in pixels. */
	width: number;
	/** Height of the node in pixels. */
	height: number;
}

/**
 * Recalculate positions for all SmartArt nodes for the current layout type;
 * `undefined` when the diagram has no content nodes. Omitted coordinates use
 * the local diagram origin, as when a host supplies only the frame's size.
 */
export function reflowSmartArtLayout(
	data: PptxSmartArtData,
	bounds: Pick<DiagramReflowBounds, 'width' | 'height'> &
		Partial<Pick<DiagramReflowBounds, 'x' | 'y'>>,
): PptxSmartArtDrawingShape[] | undefined {
	return reflowDiagramLayout(data, {
		...bounds,
		x: bounds.x ?? 0,
		y: bounds.y ?? 0,
	});
}

/** Resolve a raw layout type string to a layout family by keyword (`list` by default). */
export function resolveLayoutCategory(layoutType: string | undefined): SmartArtLayoutType {
	return resolveDiagramReflowCategory(layoutType);
}
