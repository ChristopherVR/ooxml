// An approximate drawing for a Word SmartArt that has no cached drawing, through the shared
// `approximateDiagramDrawingFromNodes` (`diagram/approximate.ts`, which Excel uses too).
import { previewDiagramLayout, type DiagramDrawing } from '../../diagram/index';
import type { DocxDiagram, ThemeCatalog } from '../index';
import { themeDrawing } from './smartart-theme';

export function previewDocxDiagram(diagram: DocxDiagram, theme?: ThemeCatalog) {
	return previewDiagramLayout({
		extentEmu: diagram.extentEmu,
		nodes: diagram.nodes,
		layoutId: diagram.layout?.uniqueId,
		family: diagram.layout?.family,
		colorsId: diagram.colorsId,
		source: diagram.layoutSource,
		scheme: themeDrawing({ shapes: [], issues: [] }, theme).schemeColors,
		minorFont: theme?.fonts.minor.latin,
		formatScheme: theme?.formatScheme,
	});
}

/** The approximate drawing, in EMU relative to the diagram frame; undefined without nodes. */
export function approximateDiagramDrawing(diagram: DocxDiagram): DiagramDrawing | undefined {
	return previewDocxDiagram(diagram).drawing;
}
