/** Pure touch-gesture math: pinch zoom, tap/drag/long-press classification, hit tolerance. */
export interface TouchPoint {
	readonly x: number;
	readonly y: number;
}
/** Movement under this many CSS pixels still counts as a tap or a held finger. */
export const TOUCH_SLOP = 8;
/** A finger held still this long (ms) opens the context menu. */
export const LONG_PRESS_MS = 500;
/** Extra hit radius (CSS px) around thin shapes for a finger. */
export const TOUCH_HIT_RADIUS = 14;
export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 8;

export function pinchDistance(a: TouchPoint, b: TouchPoint): number {
	return Math.hypot(a.x - b.x, a.y - b.y);
}
export function pinchMidpoint(a: TouchPoint, b: TouchPoint): TouchPoint {
	return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
}
export function clampZoom(zoom: number, min = MIN_ZOOM, max = MAX_ZOOM): number {
	return Number.isFinite(zoom) ? Math.max(min, Math.min(max, zoom)) : 1;
}
/** Zoom after the fingers moved from `startDistance` apart to `distance` apart. */
export function pinchZoom(
	startZoom: number,
	startDistance: number,
	distance: number,
	min = MIN_ZOOM,
	max = MAX_ZOOM,
): number {
	if (!(startDistance > 0) || !(distance > 0)) return clampZoom(startZoom, min, max);
	return clampZoom((startZoom * distance) / startDistance, min, max);
}
export interface TouchRect {
	readonly left: number;
	readonly top: number;
	readonly width: number;
	readonly height: number;
}
/** Where a screen point sits inside a rectangle, as a 0..1 fraction on each axis. */
export function pointFraction(point: TouchPoint, rect: TouchRect): TouchPoint {
	return {
		x: rect.width > 0 ? (point.x - rect.left) / rect.width : 0.5,
		y: rect.height > 0 ? (point.y - rect.top) / rect.height : 0.5,
	};
}
/**
 * How far to scroll so the content point at `fraction` of `rect` (the paper after zooming) sits
 * under the screen point `target`. Anchoring on the paper keeps its fixed margins out of the maths.
 */
export function pinchScrollDelta(
	fraction: TouchPoint,
	rect: TouchRect,
	target: TouchPoint,
): TouchPoint {
	return {
		x: rect.left + fraction.x * rect.width - target.x,
		y: rect.top + fraction.y * rect.height - target.y,
	};
}
export type TouchPhase = 'tap' | 'long-press' | 'drag' | 'pending';
/** Tap (released quickly, still), long-press (held still), drag (moved) or pending (held, short). */
export function classifyTouch(input: {
	moved: number;
	elapsed: number;
	released: boolean;
	slop?: number;
	longPressMs?: number;
}): TouchPhase {
	const { moved, elapsed, released } = input;
	if (moved > (input.slop ?? TOUCH_SLOP)) return 'drag';
	if (elapsed >= (input.longPressMs ?? LONG_PRESS_MS)) return 'long-press';
	return released ? 'tap' : 'pending';
}
/** Hit radius in px: fingers get a wider one than a mouse or pen. */
export function hitTolerance(pointerType: string): number {
	return pointerType === 'touch' ? TOUCH_HIT_RADIUS : 0;
}
/** Probe offsets (centre first, then two rings) for finding a thin shape near a finger. */
export function hitProbes(radius: number): TouchPoint[] {
	const probes: TouchPoint[] = [{ x: 0, y: 0 }];
	if (radius <= 0) return probes;
	for (const r of [radius / 2, radius])
		for (let i = 0; i < 8; i++) {
			const angle = (i * Math.PI) / 4;
			probes.push({ x: Math.cos(angle) * r, y: Math.sin(angle) * r });
		}
	return probes;
}
