// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// SmartArt editing in a Word document: applies the shared node commands (`diagram/edit`,
// `applyDiagramNodeEdit`, which Excel uses too) to a `DocxDiagram`, and creates a new diagram from
// an Insert SmartArt preset. An edit returns a new diagram (the editor keeps the old one for undo)
// with the nodes in tree order and its cached drawing dropped: Word lays an edited diagram out
// again from its data and layout parts and ignores a stale drawing (docs/agnostic-core-plan.md),
// so the save removes it.
import {
	applyDiagramNodeEdit,
	createDiagramNodes,
	type DiagramEditOptions,
	type DiagramLayoutType,
	type DiagramNodeEdit,
} from '../diagram/index';
import { diagramNotice, type DocxDiagram, type DocxDiagramNode } from './diagram';
import { treeOrderDiagramNodes } from './diagram-nodes';

/** One SmartArt Design or text pane command (the shared `DiagramNodeEdit`). */
export type DocxDiagramEdit = DiagramNodeEdit;

export interface DocxDiagramEditResult {
	diagram: DocxDiagram;
	/** The node the command created, or the node it acted on (for selection). */
	nodeId?: string;
	/** False when the command did not apply (unknown node, first sibling demoted, ...). */
	changed: boolean;
}

/** The diagram as edited: new nodes, no cached drawing. */
export function withDiagramNodes(diagram: DocxDiagram, nodes: DocxDiagramNode[]): DocxDiagram {
	const { drawing: _drawing, ...rest } = diagram;
	const parts = { ...diagram.parts };
	const drawingRelId = parts.drawing?.relId;
	delete parts.drawing;
	const hostRelationships = { ...diagram.hostRelationships };
	if (drawingRelId) delete hostRelationships[drawingRelId];
	return {
		...rest,
		parts,
		hostRelationships,
		nodes: treeOrderDiagramNodes(nodes),
		rendering: 'placeholder',
		notice: diagramNotice('placeholder'),
		issues: diagram.issues.filter((issue) => !issue.code.startsWith('DIAGRAM_DRAWING')),
	};
}

/**
 * Applies one command. Returns the diagram unchanged (`changed: false`) when the command does not
 * apply; otherwise a new diagram with the edited nodes and the cached drawing dropped.
 */
export function editDocxDiagram(
	diagram: DocxDiagram,
	edit: DocxDiagramEdit,
	options?: DiagramEditOptions,
): DocxDiagramEditResult {
	const result = applyDiagramNodeEdit(diagram.nodes, edit, options);
	const nodeId = result.nodeId ? { nodeId: result.nodeId } : {};
	if (!result.changed) return { diagram, changed: false, ...nodeId };
	return { diagram: withDiagramNodes(diagram, result.nodes), changed: true, ...nodeId };
}

/**
 * The diagram laid out with another layout (SmartArt Design > Change Layout). The caller supplies
 * the new definition as pending media under the diagram's layout part name; the save writes it,
 * names it on the data part and drops the presentation points laid out for the old layout.
 */
export function changeDocxDiagramLayout(
	diagram: DocxDiagram,
	layoutId: string,
	family?: DiagramLayoutType,
): DocxDiagram {
	if (diagram.layout?.uniqueId === layoutId) return diagram;
	return {
		...withDiagramNodes(diagram, diagram.nodes),
		layout: { uniqueId: layoutId, ...(family ? { family } : {}) },
	};
}

/** What Insert SmartArt needs to create a diagram. */
export interface DocxDiagramCreation {
	/** `uniqueId` of the layout definition supplied as the layout part. */
	layoutId: string;
	/** Category of the layout (`list`, `process`...), for the approximate drawing. */
	family?: DiagramLayoutType;
	colorsId: string;
	quickStyleId: string;
	/** The node texts the diagram starts with (top level, in order). */
	items: readonly string[];
	widthPx: number;
	heightPx: number;
	/**
	 * Unique suffix of the new part names (`word/diagrams/data<suffix>.xml`...). The caller
	 * supplies the layout, colour and quick-style parts under these names as pending media.
	 */
	partSuffix: string;
	/** `wp:docPr/@name`; `Diagram` by default. */
	name?: string;
	createModelId?: () => string;
}

const EMU_PER_PX = 9525;

/** The four part names a new diagram is saved under. */
export function insertedDiagramPartNames(
	suffix: string,
): Record<'data' | 'layout' | 'quickStyle' | 'colors', string> {
	return {
		data: `word/diagrams/data${suffix}.xml`,
		layout: `word/diagrams/layout${suffix}.xml`,
		quickStyle: `word/diagrams/quickStyle${suffix}.xml`,
		colors: `word/diagrams/colors${suffix}.xml`,
	};
}

/** A new SmartArt graphic, not yet in any package, with `items` as its top-level nodes. */
export function createDocxDiagram(creation: DocxDiagramCreation): DocxDiagram {
	const names = insertedDiagramPartNames(creation.partSuffix);
	const options = creation.createModelId ? { createModelId: creation.createModelId } : undefined;
	const part = (partName: string) => ({ relId: '', partName });
	return {
		placement: 'inline',
		extentEmu: {
			width: Math.round(creation.widthPx * EMU_PER_PX),
			height: Math.round(creation.heightPx * EMU_PER_PX),
		},
		name: creation.name ?? 'Diagram',
		relationshipIds: {},
		parts: {
			data: part(names.data),
			layout: part(names.layout),
			quickStyle: part(names.quickStyle),
			colors: part(names.colors),
		},
		hostRelationships: {},
		layout: {
			uniqueId: creation.layoutId,
			...(creation.family ? { family: creation.family } : {}),
		},
		colorsId: creation.colorsId,
		quickStyleId: creation.quickStyleId,
		// A hierarchy starts as one top node over the others, as Word's own Insert does.
		nodes: createDiagramNodes(creation.items, creation.family, options),
		rendering: 'placeholder',
		notice: diagramNotice('placeholder'),
		issues: [],
		inserted: true,
	};
}
