// An approximate drawing for a Word SmartArt that has no cached drawing, through the shared
// `approximateDiagramDrawingFromNodes` (`diagram/approximate.ts`, which Excel uses too).
import { approximateDiagramDrawingFromNodes, type DiagramDrawing } from '../../diagram/index';
import type { DocxDiagram } from '../index';

/** The approximate drawing, in EMU relative to the diagram frame; undefined without nodes. */
export function approximateDiagramDrawing(diagram: DocxDiagram): DiagramDrawing | undefined {
	return approximateDiagramDrawingFromNodes({
		extentEmu: diagram.extentEmu,
		nodes: diagram.nodes,
		layoutId: diagram.layout?.uniqueId,
		family: diagram.layout?.family,
		colorsId: diagram.colorsId,
	});
}
