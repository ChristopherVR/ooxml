// An approximate drawing for a SmartArt graphic that has no cached drawing: one inserted or edited
// in this editor, or a file saved without one. The shared reflow (`edit/reflow.ts`) lays the nodes
// out in the diagram's frame by layout family; this is not the layout Office computes, which it
// does again from the data and layout parts when it opens the file. Moved from
// `docx/ui/smartart-approximate.ts` so Word and Excel draw such a diagram the same way.
import { diagramDrawingFromViews } from './drawing-from-views';
import { reflowDiagramLayout } from './edit/reflow';
import type { DiagramNode } from './model/node';
import type { DiagramDrawing, DiagramLayoutType } from './types';

const EMU_PER_PX = 9525;
const COLORFUL = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

export interface ApproximateDiagramInput {
	/** The frame in EMU; the drawing is relative to it. */
	extentEmu: { width: number; height: number };
	nodes: readonly { id: string; text: string; parentId?: string | undefined }[];
	layoutId?: string | undefined;
	family?: DiagramLayoutType | undefined;
	colorsId?: string | undefined;
}

/** The approximate drawing, in EMU relative to the diagram frame; undefined without nodes. */
export function approximateDiagramDrawingFromNodes(
	input: ApproximateDiagramInput,
): DiagramDrawing | undefined {
	const width = input.extentEmu.width / EMU_PER_PX;
	const height = input.extentEmu.height / EMU_PER_PX;
	if (!(width > 0 && height > 0) || input.nodes.length === 0) return undefined;
	// An empty node is still a shape (Office shows its "[Text]" placeholder); the reflow skips
	// nodes without text, so it gets a space.
	const nodes: DiagramNode[] = input.nodes.map((node) => ({
		id: node.id,
		text: node.text || ' ',
		...(node.parentId ? { parentId: node.parentId } : {}),
	}));
	const views = reflowDiagramLayout(
		{
			nodes,
			...(input.family ? { resolvedLayoutType: input.family } : {}),
			...(input.layoutId ? { layoutType: input.layoutId } : {}),
		},
		{ x: 0, y: 0, width, height },
	);
	if (!views?.length) return undefined;
	return diagramDrawingFromViews(views, {
		accents: input.colorsId?.includes('colorful') ? COLORFUL : ['accent1'],
	});
}
