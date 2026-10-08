/**
 * OOXML preset bend geometry for multi-segment elbow connectors
 * (`bentConnector3/4/5`, `curvedConnector3/4/5`).
 *
 * PowerPoint's elbow connectors do NOT avoid obstacles (obstacle-avoiding A*
 * routing lives in `connector-router.ts` and is applied separately by
 * `connector-path.ts`, only when a binding supplies an obstacle list; that is
 * out of scope here). The ECMA-376 `bentConnector3/4/5` and
 * `curvedConnector3/4/5` paths are always authored horizontal-first against
 * their local box. `a:xfrm/@flipH` and `@flipV` orient that stored path; the
 * bounding-box aspect ratio does not transpose its axes.
 *
 * The geometry lives in `ooxml-core/geometry`; this module keeps the
 * pptx-typed adjustment helpers and `{ x, y }` adapters. No framework imports.
 */

import {
	curvedElbowPathD,
	elbowSegmentCount,
	elbowWaypoints as coreElbowWaypoints,
	type ElbowSegments,
} from 'ooxml-core/geometry';
import type { PptxElement } from 'ooxml-core/pptx';

import type { RouterPoint } from './connector-router-types';

/**
 * Normalise one of a connector's OOXML adjustment values (`adj1`/`adj2`/`adj3`,
 * falling back to the generic `adj`) to a 0..1 fraction that positions an
 * elbow bend line or curve control point. OOXML stores these in 1000ths of a
 * percent (0..100000); values already in 0..1 are passed through. Defaults to
 * `fallback` (the spec midpoint, `0.5`) when no usable adjustment is present,
 * so an explicitly authored `adj1`/`adj2`/`adj3` always wins over the
 * auto-computed default.
 */
export function connectorAdjustmentFraction(
	element: PptxElement,
	key: string,
	fallback = 0.5,
): number {
	const adj = (element as { shapeAdjustments?: Record<string, number> }).shapeAdjustments;
	const raw = adj?.[key] ?? adj?.adj;
	if (typeof raw !== 'number' || !Number.isFinite(raw)) {
		return fallback;
	}
	const fraction = Math.abs(raw) > 1 ? raw / 100000 : raw;
	return Math.min(1, Math.max(0, fraction));
}

/**
 * Normalise a connector's first adjustment value (`adj1`/`adj`) to a 0..1
 * fraction. Kept as a named entry point for `adj1` specifically (the only
 * adjustment a `bentConnector3`/`curvedConnector3` elbow uses); see
 * {@link connectorAdjustmentFraction} for `adj2`/`adj3`.
 */
export function connectorBendFraction(element: PptxElement): number {
	return connectorAdjustmentFraction(element, 'adj1', 0.5);
}

export { curvedElbowPathD, elbowSegmentCount };
export type { ElbowSegments };

/** OOXML connector preset paths always use x as their primary bend axis. */
export function isHorizontalPrimary(_x1: number, _y1: number, _x2: number, _y2: number): boolean {
	return true;
}

/** Bend waypoints (endpoints included) as `{ x, y }` points; see `ooxml-core/geometry`. */
export function elbowWaypoints(
	x1: number,
	y1: number,
	x2: number,
	y2: number,
	segments: ElbowSegments,
	adj1: number,
	adj2: number,
	adj3: number,
): RouterPoint[] {
	return coreElbowWaypoints(
		x1,
		y1,
		x2,
		y2,
		Math.abs(x2 - x1),
		Math.abs(y2 - y1),
		segments,
		adj1,
		adj2,
		adj3,
	).map(([x, y]) => ({ x, y }));
}
