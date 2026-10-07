import type { DiagramDrawingShape } from './types.js';

/** Axis-aligned drawing bounds, in the same units as the supplied frames. */
export interface DiagramDrawingBounds {
	minX: number;
	minY: number;
	width: number;
	height: number;
}

/**
 * Fit cached shape and independent text rectangles. Rotation, stroke and
 * effects are not included. Empty or zero-area spans retain the unit fallback
 * used by the PowerPoint renderer.
 */
export function computeDiagramDrawingBounds(
	shapes: Iterable<Pick<DiagramDrawingShape, 'frame' | 'textFrame'>>,
): DiagramDrawingBounds {
	let minX = Infinity;
	let minY = Infinity;
	let maxX = -Infinity;
	let maxY = -Infinity;
	for (const { frame, textFrame = frame } of shapes) {
		const shapeMinX = Math.min(frame.x, textFrame.x);
		const shapeMinY = Math.min(frame.y, textFrame.y);
		const shapeMaxX = Math.max(frame.x + frame.width, textFrame.x + textFrame.width);
		const shapeMaxY = Math.max(frame.y + frame.height, textFrame.y + textFrame.height);
		if (shapeMinX < minX) minX = shapeMinX;
		if (shapeMinY < minY) minY = shapeMinY;
		if (shapeMaxX > maxX) maxX = shapeMaxX;
		if (shapeMaxY > maxY) maxY = shapeMaxY;
	}
	if (!Number.isFinite(minX)) return { minX: 0, minY: 0, width: 1, height: 1 };
	return { minX, minY, width: maxX - minX || 1, height: maxY - minY || 1 };
}
