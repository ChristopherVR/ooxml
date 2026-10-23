// The data part of a SmartArt graphic created by Insert SmartArt (moved from
// `docx/diagram-save.ts` so Word and Excel write it the same way): a `doc` point naming the
// layout, colour and quick-style definitions, and the nodes with their transition pairs and
// `parOf` connections. No presentation points and no cached drawing: Word and Excel lay a diagram
// out again from its data and layout parts (docs/agnostic-core-plan.md).
import { createOfficeGuid } from '../../crypto/uuid';
import { writeFabricatedDiagramData } from './data';
import { writeEditedDiagramData, type DiagramDataEditNode } from './data-edit';

export interface InsertedDiagramData {
	nodes: readonly DiagramDataEditNode[];
	layoutId: string;
	/** Layout category (`list`, `process`...); `list` by default. */
	layoutCategory?: string | undefined;
	colorsId: string;
	/** Colour category (`accent1`, `colorful`...); `accent1` by default. */
	colorsCategory?: string | undefined;
	quickStyleId: string;
}

/** The data part XML of a new diagram. */
export function writeInsertedDiagramData(
	diagram: InsertedDiagramData,
	createModelId: () => string = createOfficeGuid,
): string {
	const skeleton = writeFabricatedDiagramData(
		{ nodes: [] },
		{
			layoutUniqueId: diagram.layoutId,
			layoutCategory: diagram.layoutCategory ?? 'list',
			quickStyleUniqueId: diagram.quickStyleId,
			colorsUniqueId: diagram.colorsId,
			colorsCategory: diagram.colorsCategory ?? 'accent1',
		},
		{ createModelId },
	);
	const nodes = diagram.nodes.map(({ id, text, parentId, nodeType }) => ({
		id,
		text,
		parentId,
		nodeType,
	}));
	return writeEditedDiagramData(skeleton, nodes, { createModelId });
}
