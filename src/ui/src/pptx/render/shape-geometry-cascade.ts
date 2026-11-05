/**
 * `shape-geometry-cascade`: the single source of truth for HOW a shape's
 * geometry is painted.
 *
 * Every binding resolves fill, stroke, effects and 3D through shared modules
 * already, then ends with the same ~40-line decision tree: connector ->
 * stroke-only preset -> roundRect -> ellipse -> clip-path -> line -> cylinder.
 * That tail was hand-ported five times, and the copies drifted:
 *
 *  - Angular compared `shapeType` RAW (`=== 'ellipse' || === 'circle'`) instead
 *    of normalising it, so `oval` - a preset offered in the shape picker - and
 *    any capitalised spelling missed the ellipse branch entirely.
 *  - Angular had no connector, line or cylinder branch at all, so a cylinder
 *    fell through to a clip-path instead of its `48% / 12%` radius.
 *  - Two fidelity fixes in one day (ellipse pill radius, identity-rect clip)
 *    each had to be applied four and five times respectively.
 *
 * This module returns a framework-neutral DECISION; each binding only maps that
 * decision onto its own style-object shape (`React.CSSProperties`, a kebab-case
 * record, etc). Adding a branch here reaches all five bindings at once.
 *
 * @module render/shape-geometry-cascade
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { ShapeGeometryDecision } from './shape-geometry-base';
import { resolveBaseShapeGeometry } from './shape-geometry-base';
import { usesSilhouetteEffects } from './silhouette-effects';

export type { ShapeGeometryDecision } from './shape-geometry-base';

/**
 * Resolve a shape element's geometry to a neutral painting decision.
 *
 * The base cascade (`resolveBaseShapeGeometry`) is unchanged: connectors and
 * open presets never get a box, the two cheap radius approximations win over a
 * clip-path, and `line`/`cylinder` are consulted only once the evaluator has
 * declined to produce geometry. One thing overrides it: a shape whose outer
 * shadow or glow is painted as a silhouette (`silhouette-effects`) must not be
 * clipped, since a clip-path would cut those effects off; its fill is painted by
 * the SVG overlay instead (`subpath-fill-overlay`).
 *
 * @param element The shape-like element to resolve.
 * @returns The decision a binding should apply to its style object.
 */
export function resolveShapeGeometry(element: PptxElement): ShapeGeometryDecision {
	const decision = resolveBaseShapeGeometry(element);
	if (decision.kind === 'clipPath' && usesSilhouetteEffects(element)) {
		return { kind: 'none' };
	}
	return decision;
}
