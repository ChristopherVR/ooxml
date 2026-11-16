/**
 * presentation-annotations-helpers.ts: Pure geometry helpers for presentation
 * ink annotations (pen, highlighter, eraser, laser).
 *
 * Ported from React:
 *   packages/react/src/viewer/components/PresentationAnnotationOverlay.tsx
 *   packages/react/src/viewer/hooks/usePresentationAnnotations.ts
 *
 * The stroke-path and cursor helpers live in `pptx-viewer-shared`
 * (`render/annotation-overlay.ts`, shared by React/Vue/Angular) and are
 * re-wrapped here with this file's own local point/tool types; everything
 * else (smoothing, eraser hit-testing, laser fade, stroke ids) is
 * Angular-only and stays pure so it can be unit-tested without TestBed.
 */
import { buildStrokePathD, cursorForTool as sharedCursorForTool } from 'ooxml-ui/pptx';

// Stroke widths, opacity and the eraser radius are the shared ink controller's.
export {
	PRESENTATION_ERASER_RADIUS as ERASER_RADIUS,
	PRESENTATION_HIGHLIGHTER_OPACITY as HIGHLIGHTER_OPACITY,
	PRESENTATION_HIGHLIGHTER_WIDTH as HIGHLIGHTER_WIDTH,
	PRESENTATION_PEN_WIDTH as PEN_WIDTH,
} from 'ooxml-ui/pptx';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** A single {x, y} coordinate in slide-space pixels. */
export interface AnnotationPoint {
	x: number;
	y: number;
}

/** An ink stroke: a sequence of points with visual properties. */
export interface AnnotationStroke {
	id: string;
	points: AnnotationPoint[];
	color: string;
	width: number;
	/** 1 = opaque (pen); 0.4 = semi-transparent (highlighter). */
	opacity: number;
}

/** The tool currently armed in presentation mode. */
export type PresentationTool = 'none' | 'pen' | 'highlighter' | 'eraser' | 'laser';

/** Per-slide annotation storage: slide index → strokes. */
export type SlideAnnotationMap = Map<number, AnnotationStroke[]>;

/** Transient laser-pointer position in slide-space pixels. */
export interface LaserPosition {
	x: number;
	y: number;
}

// ---------------------------------------------------------------------------
// SVG path generation
// ---------------------------------------------------------------------------

/**
 * Convert a sequence of points to an SVG path `d` attribute string using
 * M + L commands (no bezier smoothing, matches the React original).
 *
 * Returns an empty string for an empty point array.
 *
 * @example
 * buildPathD([{x:0,y:0},{x:10,y:5}]) // "M 0 0 L 10 5"
 */
export function buildPathD(points: AnnotationPoint[]): string {
	return buildStrokePathD(points);
}

// ---------------------------------------------------------------------------
// Point smoothing
// ---------------------------------------------------------------------------

/**
 * Apply a simple sliding-window average to a series of points, reducing
 * jitter without imposing bezier curve complexity.
 *
 * A window size of 1 is a no-op (identity). The first and last points are
 * always preserved as-is.
 *
 * @param points  Raw pointer-event coordinates.
 * @param window  Number of neighbours on each side to average (default 2).
 */
export function smoothPoints(points: AnnotationPoint[], window = 2): AnnotationPoint[] {
	const n = points.length;
	if (n <= 2 || window <= 0) {
		return points;
	}
	const out: AnnotationPoint[] = [];
	for (let i = 0; i < n; i++) {
		if (i === 0 || i === n - 1) {
			out.push(points[i]);
			continue;
		}
		const lo = Math.max(0, i - window);
		const hi = Math.min(n - 1, i + window);
		let sx = 0;
		let sy = 0;
		const count = hi - lo + 1;
		for (let j = lo; j <= hi; j++) {
			sx += points[j].x;
			sy += points[j].y;
		}
		out.push({ x: sx / count, y: sy / count });
	}
	return out;
}

// ---------------------------------------------------------------------------
// Laser dot fade
// ---------------------------------------------------------------------------

/**
 * Compute the CSS `opacity` value for a fading laser dot given a fade ratio
 * in [0, 1] where 0 = fully visible and 1 = fully invisible.
 *
 * Uses a simple quadratic ease-out so the dot lingers at full intensity then
 * fades quickly near the end.
 *
 * @param ratio  0 (fresh) → 1 (expired)
 */
export function laserDotOpacity(ratio: number): number {
	const clamped = Math.max(0, Math.min(1, ratio));
	return (1 - clamped) * (1 - clamped);
}

// ---------------------------------------------------------------------------
// Cursor helper
// ---------------------------------------------------------------------------

/**
 * Return the CSS `cursor` value that matches `tool`.
 */
export function cursorForTool(tool: PresentationTool): string {
	return sharedCursorForTool(tool);
}
