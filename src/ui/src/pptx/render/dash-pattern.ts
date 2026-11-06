/**
 * PowerPoint's measured dash patterns. The logic lives in the core
 * (`ooxml-core/drawingml`, `dash-pattern.ts`) so other formats can use it;
 * this module keeps the render-layer exports in place.
 *
 * @module render/dash-pattern
 */
export {
	capCompensatedDashArray,
	dashPatternUnits,
	dashPenWidth,
	isDashedStroke,
} from 'ooxml-core/drawingml';
export type { DashSegmentInput } from 'ooxml-core/drawingml';
