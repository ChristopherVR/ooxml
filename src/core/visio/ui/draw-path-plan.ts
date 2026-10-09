import type { VisioGeometryEdit } from '../edit-commands';
import type { VisioPage } from '../model';
import {
	visioCleanSamples,
	visioFitFreeform,
	visioFitPencil,
	visioPathBounds,
	visioPathSamples,
	visioQuarterArc,
	type VisioPathPoint,
	type VisioPathSegment,
} from '../path-fit';
import { visioNextShapeId } from './shape-id';
import { visioPageEditToDrawing } from './page-edit';
import { visioLineCreationCommand, type VisioDrawingPoint } from './draw-plan';

/** Path drawing tools: Freeform (Ctrl+5), Pencil (Ctrl+4) and Arc (Ctrl+7). */
export type VisioPathTool = 'freeform' | 'pencil' | 'arc';
export const isVisioPathTool = (tool: unknown): tool is VisioPathTool =>
	tool === 'freeform' || tool === 'pencil' || tool === 'arc';
/** Page inches within which a freeform or pencil stroke that returns to its start closes. */
export const VISIO_PATH_CLOSE_DISTANCE = 0.125;
/** Default fitting tolerance in page inches. */
export const VISIO_PATH_TOLERANCE = 1 / 32;
const MINIMUM = 1 / 16;
const MAX_SAMPLES = 20_000;
const bounded = (value: number) => Number.isFinite(value) && Math.abs(value) <= 1_000_000;

/** Page samples (top-left origin) to the edit API's bottom-left drawing orientation. */
const flip = (page: VisioPage, point: VisioPathPoint): VisioPathPoint => ({
	x: point.x,
	y: page.height - point.y,
});

function strokeExtent(points: readonly VisioPathPoint[]): number {
	const xs = points.map((point) => point.x),
		ys = points.map((point) => point.y);
	return Math.hypot(Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys));
}

/**
 * Plan a Freeform, Pencil or Arc creation from the sampled pointer stroke in page inches. Returns
 * undefined for a stroke too small to draw. A stroke that simplifies to one straight segment, or
 * whose path has no area extent, becomes an ordinary line (as a single Pencil segment is in Visio).
 */
export function visioPathDrawPlan(
	page: VisioPage,
	kind: VisioPathTool,
	samples: readonly VisioDrawingPoint[],
	options: { tolerance?: number } = {},
): VisioGeometryEdit | undefined {
	if (!isVisioPathTool(kind) || !Array.isArray(samples) || samples.length > MAX_SAMPLES)
		throw new Error('Invalid path drawing tool or samples.');
	const tolerance = options.tolerance ?? VISIO_PATH_TOLERANCE;
	if (!(tolerance > 0) || !bounded(tolerance)) throw new Error('Invalid path fitting tolerance.');
	const points = visioCleanSamples(samples).filter((point) => bounded(point.x) && bounded(point.y));
	if (points.length < 2 || strokeExtent(points) < MINIMUM) return undefined;
	const first = points[0]!,
		last = points[points.length - 1]!;
	const line = () => {
		// The farthest sample from the start keeps the stroke's reach when it has no area.
		let end = last;
		for (const point of points)
			if (
				Math.hypot(point.x - first.x, point.y - first.y) >
				Math.hypot(end.x - first.x, end.y - first.y)
			)
				end = point;
		return visioLineCreationCommand(page, first, end);
	};
	let segments: VisioPathSegment[];
	let closed = false;
	const drawing = points.map((point) => flip(page, point));
	if (kind === 'arc') {
		const arc = visioQuarterArc(drawing[0]!, drawing[drawing.length - 1]!);
		if (!arc) return line();
		segments = [arc];
	} else {
		closed =
			points.length >= 4 &&
			Math.hypot(last.x - first.x, last.y - first.y) <= VISIO_PATH_CLOSE_DISTANCE &&
			strokeExtent(points) > 2 * VISIO_PATH_CLOSE_DISTANCE;
		if (closed) drawing[drawing.length - 1] = { ...drawing[0]! };
		segments =
			kind === 'freeform'
				? visioFitFreeform(drawing, tolerance, closed)
				: visioFitPencil(drawing, tolerance);
		if (segments.length < (closed ? 2 : 1)) return line();
		if (!closed && segments.length === 1 && segments[0]!.kind === 'line') return line();
	}
	const start = drawing[0]!;
	const bounds = visioPathBounds(start, segments);
	if (bounds.maxX - bounds.minX < 1e-6 || bounds.maxY - bounds.minY < 1e-6) return line();
	const command = visioPageEditToDrawing(page, {
		type: 'create-path',
		pageId: page.id,
		shapeId: visioNextShapeId(page),
		x: start.x,
		y: start.y,
		segments,
		...(closed ? { closed } : {}),
	}) as VisioGeometryEdit;
	return command;
}

/** SVG path data (page inches, top-left origin) previewing a stroke while it is drawn. */
export function visioPathPreview(
	kind: VisioPathTool,
	samples: readonly VisioDrawingPoint[],
): string {
	const points = visioCleanSamples(samples).slice(0, MAX_SAMPLES);
	if (!points.length) return '';
	const outline =
		kind === 'arc' && points.length > 1
			? (() => {
					const arc = visioQuarterArc(points[0]!, points[points.length - 1]!);
					return arc
						? visioPathSamples(points[0]!, [arc], 24)
						: [points[0]!, points[points.length - 1]!];
				})()
			: points;
	return outline.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ');
}
