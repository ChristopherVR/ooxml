// An approximate drawing for a Word SmartArt that has no cached drawing: one inserted or edited in
// this editor, or a file saved without one. The shared reflow (`diagram/edit/reflow.ts`) lays the
// nodes out in the diagram's frame by layout family; this is not the layout Word computes, which
// it does again from the data and layout parts when it opens the file.
import {
	diagramDrawingFromViews,
	reflowDiagramLayout,
	type DiagramDrawing,
	type DiagramNode,
} from '../../diagram/index';
import type { DocxDiagram } from '../index';

const EMU_PER_PX = 9525;
const COLORFUL = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** The approximate drawing, in EMU relative to the diagram frame; undefined without nodes. */
export function approximateDiagramDrawing(diagram: DocxDiagram): DiagramDrawing | undefined {
	const width = diagram.extentEmu.width / EMU_PER_PX;
	const height = diagram.extentEmu.height / EMU_PER_PX;
	if (!(width > 0 && height > 0) || diagram.nodes.length === 0) return undefined;
	// An empty node is still a shape (Word shows its "[Text]" placeholder); the reflow skips
	// nodes without text, so it gets a space.
	const nodes: DiagramNode[] = diagram.nodes.map((node) => ({
		id: node.id,
		text: node.text || ' ',
		...(node.parentId ? { parentId: node.parentId } : {}),
	}));
	const views = reflowDiagramLayout(
		{
			nodes,
			...(diagram.layout?.family ? { resolvedLayoutType: diagram.layout.family } : {}),
			...(diagram.layout?.uniqueId ? { layoutType: diagram.layout.uniqueId } : {}),
		},
		{ x: 0, y: 0, width, height },
	);
	if (!views?.length) return undefined;
	return diagramDrawingFromViews(views, {
		accents: diagram.colorsId?.includes('colorful') ? COLORFUL : ['accent1'],
	});
}
