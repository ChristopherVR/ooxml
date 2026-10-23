// One SmartArt Design or text pane command applied to a host's plain node list: the command
// vocabulary Word and Excel share (moved from `docx/diagram-edit.ts`), over the node operations
// of `node-ops.ts` and `node-order.ts`. Nodes stay in tree order and keep only the fields a host
// model stores, so an edit result can go straight back into the host's undo history.
import type { DiagramLayoutType } from '../types';
import type { DiagramNeutralTextParagraph } from '../text/node-text-neutral';
import { addDiagramNodeAbove, addDiagramNodeAfter, addDiagramNodeBefore } from './node-ops';
import { addDiagramNodeBelow, removeDiagramNode, setDiagramNodeText } from './node-ops';
import { demoteDiagramNode, moveDiagramNode, promoteDiagramNode } from './node-order';
import { treeOrderDiagramNodes } from './node-tree';
import type { DiagramEditOptions } from './types';

/** A content node as a host model keeps it: array order is tree order. */
export interface DiagramHostNode {
	id: string;
	text: string;
	/** Model id of the parent content node, when the data model has one. */
	parentId?: string;
	/** The node's paragraphs on the neutral text model, kept while its text is unchanged. */
	paragraphs?: DiagramNeutralTextParagraph[];
	/** `asst` for an assistant point (org charts); absent for an ordinary node. */
	nodeType?: 'asst';
}

/** One SmartArt Design or text pane command. */
export type DiagramNodeEdit =
	| {
			kind: 'add';
			position: 'after' | 'before' | 'above' | 'below';
			nodeId?: string;
			text?: string;
	  }
	| { kind: 'remove'; nodeId: string }
	| { kind: 'text'; nodeId: string; text: string }
	| { kind: 'move'; nodeId: string; direction: 1 | -1 }
	| { kind: 'promote'; nodeId: string }
	| { kind: 'demote'; nodeId: string };

export interface DiagramNodeEditResult<N extends DiagramHostNode = DiagramHostNode> {
	nodes: N[];
	/** The node the command created, or the node it acted on (for selection). */
	nodeId?: string;
	/** False when the command did not apply (unknown node, first sibling demoted, ...). */
	changed: boolean;
}

interface EditableNodes {
	nodes: DiagramHostNode[];
}

function applyOperation(
	data: EditableNodes,
	edit: DiagramNodeEdit,
	options: DiagramEditOptions | undefined,
): EditableNodes {
	switch (edit.kind) {
		case 'add': {
			const text = edit.text ?? '';
			if (edit.position === 'below') return addDiagramNodeBelow(data, edit.nodeId, text, options);
			if (edit.position === 'above')
				return edit.nodeId ? addDiagramNodeAbove(data, edit.nodeId, text, options) : data;
			const add = edit.position === 'after' ? addDiagramNodeAfter : addDiagramNodeBefore;
			return add(data, text, edit.nodeId, options);
		}
		case 'remove':
			return removeDiagramNode(data, edit.nodeId);
		case 'text':
			return setDiagramNodeText(data, edit.nodeId, edit.text);
		case 'move':
			return moveDiagramNode(data, edit.nodeId, edit.direction);
		case 'promote':
			return promoteDiagramNode(data, edit.nodeId);
		case 'demote':
			return demoteDiagramNode(data, edit.nodeId);
	}
}

/** Plain nodes again: the operations add bookkeeping fields host models do not keep. */
export function plainDiagramNodes(
	nodes: readonly DiagramHostNode[],
	before: readonly DiagramHostNode[],
): DiagramHostNode[] {
	const previous = new Map(before.map((node) => [node.id, node]));
	return nodes.map((node) => {
		const old = previous.get(node.id);
		const kept: DiagramHostNode = {
			id: node.id,
			text: node.text,
			...(node.parentId ? { parentId: node.parentId } : {}),
			...(node.nodeType === 'asst' ? { nodeType: 'asst' as const } : {}),
		};
		// Rich paragraphs stay only while the text they spell is unchanged.
		if (old?.paragraphs && old.text === node.text) kept.paragraphs = old.paragraphs;
		return kept;
	});
}

/**
 * Applies one command to `nodes`. Returns them unchanged (`changed: false`) when the command does
 * not apply; otherwise the edited nodes in tree order.
 */
export function applyDiagramNodeEdit(
	nodes: readonly DiagramHostNode[],
	edit: DiagramNodeEdit,
	options?: DiagramEditOptions,
): DiagramNodeEditResult {
	const after = applyOperation({ nodes: [...nodes] }, edit, options).nodes;
	const edited = plainDiagramNodes(after, nodes);
	if (JSON.stringify(edited) === JSON.stringify(plainDiagramNodes(nodes, nodes)))
		return {
			nodes: [...nodes],
			changed: false,
			...('nodeId' in edit && edit.nodeId ? { nodeId: edit.nodeId } : {}),
		};
	const known = new Set(nodes.map((node) => node.id));
	const created = edited.find((node) => !known.has(node.id))?.id;
	const nodeId = created ?? (edit.kind === 'remove' ? undefined : edit.nodeId);
	return {
		nodes: treeOrderDiagramNodes(edited),
		changed: true,
		...(nodeId ? { nodeId } : {}),
	};
}

/**
 * The nodes of a new diagram (Insert SmartArt) with `items` as their texts: top-level nodes in
 * order, or for a hierarchy one top node over the others, as Office's own Insert does.
 */
export function createDiagramNodes(
	items: readonly string[],
	family?: DiagramLayoutType,
	options?: DiagramEditOptions,
): DiagramHostNode[] {
	let data: EditableNodes = { nodes: [] };
	const hierarchy = family === 'hierarchy';
	for (const text of items)
		data = addDiagramNodeBelow(data, hierarchy ? data.nodes[0]?.id : undefined, text, options);
	return plainDiagramNodes(data.nodes, []);
}

/** One text pane line: a node's text and its depth below the top level. */
export interface DiagramOutlineLine {
	id: string;
	text: string;
	level: number;
}

/** The nodes (in tree order) as text pane lines. */
export function diagramNodeOutline(nodes: readonly DiagramHostNode[]): DiagramOutlineLine[] {
	const ids = new Set(nodes.map((node) => node.id));
	const levels = new Map<string, number>();
	return nodes.map((node) => {
		const parent = node.parentId && ids.has(node.parentId) ? node.parentId : undefined;
		const level = parent === undefined ? 0 : (levels.get(parent) ?? 0) + 1;
		levels.set(node.id, level);
		return { id: node.id, text: node.text, level };
	});
}
