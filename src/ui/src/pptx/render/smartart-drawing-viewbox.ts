/** PowerPoint adapter for the format-neutral cached drawing bounds. */
import { computeDiagramDrawingBounds } from 'ooxml-core/diagram';
import type { DiagramDrawingBounds } from 'ooxml-core/diagram';
import type { PptxSmartArtDrawingShape } from 'ooxml-core/pptx';

export type DrawingViewBox = DiagramDrawingBounds;

/** Preserve the per-field fallback of PowerPoint's independent text rectangle. */
export function computeDrawingViewBox(shapes: readonly PptxSmartArtDrawingShape[]): DrawingViewBox {
	return computeDiagramDrawingBounds(
		shapes.map((shape) => ({
			frame: shape,
			textFrame: {
				x: shape.textFrameX ?? shape.x,
				y: shape.textFrameY ?? shape.y,
				width: shape.textFrameWidth ?? shape.width,
				height: shape.textFrameHeight ?? shape.height,
			},
		})),
	);
}
