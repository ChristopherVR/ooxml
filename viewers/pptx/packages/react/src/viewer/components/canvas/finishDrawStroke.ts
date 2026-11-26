import { buildFreeformShapeElement, strokeToInkElement } from 'ooxml-ui/pptx';
import type { InkPoint } from 'ooxml-ui/pptx';
import type { InkPptxElement, ShapePptxElement } from 'pptx-viewer-core';

import type { DrawingTool } from '../../types-ui';

/** The accumulated Draw-tool gesture state `finishDrawStroke` turns into an element. */
export interface FinishDrawStrokeInput {
	tool: DrawingTool;
	points: InkPoint[];
	pressures: number[];
	tiltX: number[];
	tiltY: number[];
	color: string;
	width: number;
}

/** The element a finished stroke commits, tagged by which kind it is. */
export type FinishedDrawStroke =
	| { kind: 'ink'; element: InkPptxElement }
	| { kind: 'freeform'; element: ShapePptxElement };

/**
 * Turn an accumulated Draw-tool stroke into a committed `ink` element or
 * `freeform` shape, or `null` for a too-short stroke (a plain tap).
 *
 * Split out of `useDrawingOverlay`'s `handleDrawPointerUp` (which was pushing
 * that file past this repo's 300-LOC guideline): pure geometry/element
 * construction with no React dependency, so it is unit-testable on its own.
 */
export function finishDrawStroke(input: FinishDrawStrokeInput): FinishedDrawStroke | null {
	const { tool, points, pressures, tiltX, tiltY, color, width } = input;
	if (points.length < 2) {
		return null;
	}
	if (tool === 'freeform') {
		// The shared builder pads the box by the stroke width exactly as the ink
		// path below does, and closes the path only when the stroke ends back on
		// its start point (PowerPoint's Freeform/Scribble leaves it open otherwise).
		const element = buildFreeformShapeElement(points, { color, width });
		return element ? { kind: 'freeform', element } : null;
	}

	if (tool === 'select' || tool === 'eraser') {
		return null;
	}
	const element = strokeToInkElement({
		tool,
		color,
		width,
		points: points.map((point, index) => ({
			...point,
			pressure: pressures[index],
			tiltX: tiltX[index],
			tiltY: tiltY[index],
		})),
	});
	return element ? { kind: 'ink', element } : null;
}
