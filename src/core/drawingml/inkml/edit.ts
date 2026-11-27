import { pointInPolygon } from '../../geometry/shape-boolean-types';
import { inkDistance, MAX_INK_SAMPLES, type InkCoordinate } from '../ink-geometry';
import { MAX_INK_PART_SAMPLES, MAX_INK_TRACES } from './inkml-limits';
import { inkPathPoints } from './path';
import { inkPenOpacity, type InkPen } from '../ink-pens';
import type { InkStroke } from './types';

/**
 * Format-neutral edits on the strokes of one ink drawing (coordinates in CSS pixels of its box):
 * a stroke from captured pointer samples, bounds, moving and scaling, the stroke eraser and the
 * lasso. Word and Excel keep the strokes of each drawing in this form.
 *
 * @module drawingml/inkml/edit
 */

/** One captured pointer sample. */
export interface InkSample extends InkCoordinate {
	/** 0..1; mice report a constant 0.5 (or 0 while hovering), which carries no information. */
	pressure?: number;
	/** Milliseconds (any epoch), for Ink Replay. */
	time?: number;
}

export { inkPathPoints };

const round = (value: number): number => Math.round(value * 100) / 100;

function pathOf(points: readonly InkCoordinate[]): string {
	return points
		.map((point, index) => `${index ? 'L' : 'M'} ${round(point.x)} ${round(point.y)}`)
		.join(' ');
}

/** Drops samples closer than half a pixel to the last kept one: they add bytes, not shape. */
function thin(samples: readonly InkSample[]): InkSample[] {
	const kept: InkSample[] = [];
	for (const sample of samples.slice(0, MAX_INK_SAMPLES)) {
		if (!Number.isFinite(sample.x) || !Number.isFinite(sample.y)) continue;
		const last = kept.at(-1);
		if (!last || inkDistance(last, sample) >= 0.5) kept.push(sample);
	}
	const end = samples.at(-1);
	if (end && kept.length > 1 && kept.at(-1) !== end && Number.isFinite(end.x)) kept.push(end);
	return kept;
}

/**
 * A stroke from captured samples drawn with `pen`. One sample makes a dot; pressure is kept only
 * when it varies (a mouse reports a constant), and sample times feed Ink Replay.
 */
export function inkStrokeFromSamples(
	samples: readonly InkSample[],
	pen: InkPen,
): InkStroke | undefined {
	const points = thin(samples);
	const first = points[0];
	if (!first) return undefined;
	if (points.length === 1) points.push({ ...first, x: first.x + 0.01 });
	const pressures = points.map((point) => point.pressure ?? 0.5);
	const varies = pressures.some((value) => Math.abs(value - (pressures[0] ?? 0.5)) > 0.01);
	const times = points.map((point) => point.time);
	const timed = times.every((time): time is number => typeof time === 'number');
	const start = timed ? times[0] : undefined;
	const end = timed ? times.at(-1) : undefined;
	return {
		path: pathOf(points),
		color: pen.color,
		width: pen.width,
		opacity: inkPenOpacity(pen.kind),
		tool: pen.kind,
		...(varies ? { pressures: pressures.map((value) => Math.max(0, Math.min(1, value))) } : {}),
		...(start !== undefined && end !== undefined
			? {
					startTimestamp: Math.round(start),
					durationMs: Math.max(0, Math.round(end - start)),
					pointTimestamps: (times as number[]).map(Math.round),
				}
			: {}),
	};
}

export interface InkBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** The box the strokes cover, half a stroke width beyond their centre lines. */
export function inkStrokesBounds(strokes: readonly InkStroke[]): InkBox | undefined {
	let left = Infinity;
	let top = Infinity;
	let right = -Infinity;
	let bottom = -Infinity;
	for (const stroke of strokes) {
		const half = stroke.width / 2;
		for (const point of inkPathPoints(stroke.path)) {
			left = Math.min(left, point.x - half);
			top = Math.min(top, point.y - half);
			right = Math.max(right, point.x + half);
			bottom = Math.max(bottom, point.y + half);
		}
	}
	return Number.isFinite(left)
		? { x: left, y: top, width: Math.max(1, right - left), height: Math.max(1, bottom - top) }
		: undefined;
}

/** The strokes moved by `dx`, `dy` and then scaled about the origin by `sx`, `sy`. */
export function transformInkStrokes(
	strokes: readonly InkStroke[],
	dx: number,
	dy: number,
	sx = 1,
	sy = 1,
): InkStroke[] {
	return strokes.map((stroke) => ({
		...stroke,
		path: pathOf(
			inkPathPoints(stroke.path).map((point) => ({
				x: (point.x + dx) * sx,
				y: (point.y + dy) * sy,
			})),
		),
	}));
}

/** Distance from `point` to the segment `a`-`b`. */
function segmentDistance(point: InkCoordinate, a: InkCoordinate, b: InkCoordinate): number {
	const dx = b.x - a.x;
	const dy = b.y - a.y;
	const length = dx * dx + dy * dy;
	const t = length
		? Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / length))
		: 0;
	return Math.hypot(point.x - a.x - t * dx, point.y - a.y - t * dy);
}

/** Whether the eraser at `point` (radius in pixels) touches the stroke. */
export function inkStrokeHit(stroke: InkStroke, point: InkCoordinate, radius: number): boolean {
	const points = inkPathPoints(stroke.path);
	const reach = radius + stroke.width / 2;
	if (points.length === 1) return inkDistance(points[0]!, point) <= reach;
	for (let index = 1; index < points.length; index++)
		if (segmentDistance(point, points[index - 1]!, points[index]!) <= reach) return true;
	return false;
}

/** The stroke eraser: every stroke the eraser touches goes, as Office's Eraser (Stroke). */
export function eraseInkStrokes(
	strokes: readonly InkStroke[],
	point: InkCoordinate,
	radius: number,
): InkStroke[] {
	return strokes.filter((stroke) => !inkStrokeHit(stroke, point, radius));
}

/** The indexes of the strokes the lasso holds: most of whose points lie inside `polygon`. */
export function lassoInkStrokes(
	strokes: readonly InkStroke[],
	polygon: readonly InkCoordinate[],
): number[] {
	if (polygon.length < 3) return [];
	const ring = polygon.map((point) => ({ x: point.x, y: point.y }));
	return strokes.flatMap((stroke, index) => {
		const points = inkPathPoints(stroke.path);
		const inside = points.filter((point) => pointInPolygon(point, ring)).length;
		return points.length && inside / points.length >= 0.6 ? [index] : [];
	});
}

/** Strokes one drawing collects before a new stroke starts another: well inside what a part may hold. */
const JOIN_MAX_STROKES = 64;
const JOIN_MAX_SAMPLES = 16_000;

/**
 * Whether `added` may join the drawing holding `strokes` (strokes drawn one after another, as
 * Office collects them). The total stays well inside what the reader reads back from one part
 * (`MAX_INK_TRACES`, `MAX_INK_PART_SAMPLES`), so a drawing this library writes is always read in
 * full, and rewriting a growing drawing stays cheap.
 */
export function canJoinInk(strokes: readonly InkStroke[], added: readonly InkStroke[]): boolean {
	const count = strokes.length + added.length;
	if (count > Math.min(JOIN_MAX_STROKES, MAX_INK_TRACES)) return false;
	let samples = 0;
	for (const stroke of [...strokes, ...added]) samples += inkPathPoints(stroke.path).length;
	return samples <= Math.min(JOIN_MAX_SAMPLES, MAX_INK_PART_SAMPLES);
}
