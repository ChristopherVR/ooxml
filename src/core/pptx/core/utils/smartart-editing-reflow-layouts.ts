/**
 * SmartArt reflow layout implementations.
 *
 * Contains the individual reflow layout algorithms for specific SmartArt
 * layout types (cycle, matrix, pyramid, funnel, target, gear, venn,
 * timeline, relationship, chevron, bending).
 *
 * The implementations live in `diagram/edit` (`reflow-geometric`,
 * `reflow-stacked`, `reflow-directional`); this module re-exports them.
 *
 * These are "visual reflow" layouts used after editing operations to
 * reposition nodes without requiring the full layout engine.
 *
 * @module smartart-editing-reflow-layouts
 */

// ── Geometric / spatial layouts ──────────────────────────────────────────

export {
	reflowCycle,
	reflowMatrix,
	reflowTarget,
	reflowGear,
	reflowVenn,
} from '../../../diagram/edit/reflow-geometric';

// ── Stacked layouts (vertical bands with varying widths) ─────────────────

export { reflowPyramid, reflowFunnel } from '../../../diagram/edit/reflow-stacked';

// ── Directional / flow layouts ───────────────────────────────────────────

export {
	reflowTimeline,
	reflowRelationship,
	reflowChevron,
	reflowBending,
} from '../../../diagram/edit/reflow-directional';
