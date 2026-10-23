/**
 * SmartArt node creation, removal and text edits over the neutral data model.
 *
 * Every operation is pure: it returns a new data object of the caller's own
 * type (pptx: `PptxSmartArtData`) with the node list and `parOf` connections
 * updated, `drawingDirty` set, and the drawing shapes either patched (text) or
 * marked stale so the renderer falls back to the layout engine. Moved from
 * `pptx/core/utils/smartart-editing-node-ops.ts` and made generic; the model id
 * generator is injected (`DiagramEditOptions.createModelId`).
 *
 * @module diagram/edit/node-ops
 */

import { createOfficeGuid } from '../../crypto/uuid';
import type { DiagramConnection } from '../types';
import type { DiagramNode } from '../model/node';
import { markDiagramShapesStale, patchDiagramShape } from './shape-match';
import type { DiagramEditOptions, DiagramEditableData } from './types';

/** The new point's model id, through the host's generator when it supplies one. */
export const diagramModelId = (options?: DiagramEditOptions): string =>
	(options?.createModelId ?? createOfficeGuid)();

/**
 * The id nodes use as `parentId` when they have no visible parent within the
 * array -- i.e. the diagram's root/doc point id. Existing top-level nodes
 * (parsed from a real file) already carry this as their own `parentId`; a
 * brand-new top-level node needs the same value so the save pipeline can
 * anchor it under the same root when writing the diagram XML back out.
 * Returns `undefined` when the diagram has no nodes yet to infer it from.
 */
export function findDiagramRootParentId(nodes: readonly DiagramNode[]): string | undefined {
	const nodeIds = new Set(nodes.map((n) => n.id));
	for (const node of nodes) {
		if (node.parentId && !nodeIds.has(node.parentId)) {
			return node.parentId;
		}
	}
	return undefined;
}

/** Append a `parOf` connection from `parentId` to `childId` after the parent's other children. */
export function appendParentConnection(
	connections: DiagramConnection[],
	parentId: string,
	childId: string,
): void {
	const maxSrcOrd = connections
		.filter((c) => c.sourceId === parentId)
		.reduce((max, c) => Math.max(max, c.srcOrd ?? 0), -1);
	connections.push({
		sourceId: parentId,
		destId: childId,
		type: 'parOf',
		srcOrd: maxSrcOrd + 1,
		destOrd: 0,
	});
}

/** The edited data: new nodes (and connections), the drawing dirty and its shapes stale. */
export function withEditedStructure<D extends DiagramEditableData>(
	data: D,
	nodes: D['nodes'],
	connections?: DiagramConnection[],
): D {
	return {
		...data,
		nodes,
		...(connections ? { connections: connections.length > 0 ? connections : undefined } : {}),
		drawingDirty: true,
		drawingShapes: markDiagramShapesStale(data.drawingShapes),
	};
}

/**
 * Add a new node after a given sibling (PowerPoint's Add Shape After). If
 * `afterNodeId` is undefined or unknown, the node is appended at the end.
 */
export function addDiagramNodeAfter<D extends DiagramEditableData>(
	data: D,
	text: string,
	afterNodeId?: string,
	options?: DiagramEditOptions,
): D {
	return insertSibling(data, text, afterNodeId, 1, options);
}

/**
 * Add a new node before a given sibling (PowerPoint's Add Shape Before). If
 * `beforeNodeId` is undefined or unknown, the node is appended at the end.
 */
export function addDiagramNodeBefore<D extends DiagramEditableData>(
	data: D,
	text: string,
	beforeNodeId?: string,
	options?: DiagramEditOptions,
): D {
	return insertSibling(data, text, beforeNodeId, 0, options);
}

function insertSibling<D extends DiagramEditableData>(
	data: D,
	text: string,
	siblingId: string | undefined,
	offset: 0 | 1,
	options: DiagramEditOptions | undefined,
): D {
	const newId = diagramModelId(options);

	// Determine parent from the sibling node, falling back to the diagram's
	// root id for a genuinely top-level addition (no sibling to anchor to).
	let parentId: string | undefined;
	if (siblingId) {
		const sibling = data.nodes.find((n) => n.id === siblingId);
		parentId = sibling?.parentId;
	}
	parentId ??= findDiagramRootParentId(data.nodes);

	const newNode = { id: newId, text, parentId } as D['nodes'][number];

	// Insert next to the specified sibling, or at the end
	const nodes = [...data.nodes];
	const siblingIndex = siblingId ? nodes.findIndex((n) => n.id === siblingId) : -1;
	if (siblingIndex >= 0) {
		nodes.splice(siblingIndex + offset, 0, newNode);
	} else {
		nodes.push(newNode);
	}

	// Add a connection from parent to the new node
	const connections = [...(data.connections ?? [])];
	if (parentId) {
		appendParentConnection(connections, parentId, newId);
	}
	return withEditedStructure(data, nodes, connections);
}

/**
 * Add a new node as a child of a given parent (PowerPoint's Add Shape Below).
 *
 * If `parentId` is undefined, the node is added as a root-level item, using
 * the diagram's own root/doc id (inferred from an existing top-level node)
 * so it ends up as a sibling of the other top-level items rather than an
 * unparented node the save pipeline can't anchor into the diagram XML.
 * If `text` is undefined, a default label is generated.
 */
export function addDiagramNodeBelow<D extends DiagramEditableData>(
	data: D,
	parentId?: string,
	text?: string,
	options?: DiagramEditOptions,
): D {
	const resolvedParentId = parentId ?? findDiagramRootParentId(data.nodes);
	const newId = diagramModelId(options);
	const label = text ?? `Item ${data.nodes.length + 1}`;
	const newNode = { id: newId, text: label, parentId: resolvedParentId } as D['nodes'][number];
	const nodes = [...data.nodes, newNode];

	// Add a connection from parent to the new node
	const connections = [...(data.connections ?? [])];
	if (resolvedParentId) {
		appendParentConnection(connections, resolvedParentId, newId);
	}
	return withEditedStructure(data, nodes, connections);
}

/**
 * Add a new node above a given node (PowerPoint's Add Shape Above): the new
 * node takes the target's place under the target's parent, and the target
 * becomes its only child. Returns `data` unchanged for an unknown node.
 */
export function addDiagramNodeAbove<D extends DiagramEditableData>(
	data: D,
	nodeId: string,
	text: string,
	options?: DiagramEditOptions,
): D {
	const target = data.nodes.find((n) => n.id === nodeId);
	if (!target) {
		return data;
	}
	const newId = diagramModelId(options);
	const parentId = target.parentId ?? findDiagramRootParentId(data.nodes);
	const newNode = { id: newId, text, parentId } as D['nodes'][number];
	const nodes: D['nodes'] = [];
	for (const node of data.nodes) {
		if (node.id === nodeId) {
			nodes.push(newNode, { ...node, parentId: newId });
		} else {
			nodes.push(node);
		}
	}

	// The new node inherits the target's place among its parent's children.
	const connections: DiagramConnection[] = [];
	for (const connection of data.connections ?? []) {
		const isTargetEdge =
			connection.destId === nodeId &&
			(connection.type ?? 'parOf') === 'parOf' &&
			connection.sourceId === target.parentId;
		connections.push(isTargetEdge ? { ...connection, destId: newId } : { ...connection });
	}
	if (parentId && !connections.some((c) => c.destId === newId)) {
		appendParentConnection(connections, parentId, newId);
	}
	connections.push({ sourceId: newId, destId: nodeId, type: 'parOf', srcOrd: 0, destOrd: 0 });
	return withEditedStructure(data, nodes, connections);
}

/**
 * Remove a node by ID. Its children move up to the removed node's parent, the
 * connections referencing it are dropped, and drawing shapes are marked stale.
 */
export function removeDiagramNode<D extends DiagramEditableData>(data: D, nodeId: string): D {
	const removedNode = data.nodes.find((n) => n.id === nodeId);

	// Identify children of the removed node BEFORE mutating any objects
	const childIds = data.nodes.filter((n) => n.parentId === nodeId).map((n) => n.id);

	// Clone remaining nodes and re-parent children of the removed node
	const nodes = data.nodes
		.filter((n) => n.id !== nodeId)
		.map((n) => {
			if (n.parentId === nodeId) {
				// Re-parent to the removed node's parent, or promote to root
				return { ...n, parentId: removedNode?.parentId };
			}
			return { ...n };
		});

	// Remove connections referencing the deleted node and re-wire children
	const connections = (data.connections ?? [])
		.filter((c) => c.sourceId !== nodeId && c.destId !== nodeId)
		.map((c) => ({ ...c }));

	// Add connections from the removed node's parent to its children
	const grandParentId = removedNode?.parentId;
	if (grandParentId && childIds.length > 0) {
		for (const childId of childIds) {
			appendParentConnection(connections, grandParentId, childId);
		}
	}
	return withEditedStructure(data, nodes, connections);
}

/**
 * Update the text of a node by ID.
 *
 * When drawing shapes exist, updates the matching shape's text in-place so the
 * renderer path does not change (avoids polygon-to-chevron downgrade that makes
 * a pyramid look like stacked bars). When no drawing shapes are present the
 * family renderer reads text directly from the node array.
 */
export function setDiagramNodeText<D extends DiagramEditableData>(
	data: D,
	nodeId: string,
	newText: string,
): D {
	const nodes = data.nodes.map((n) => (n.id === nodeId ? { ...n, text: newText } : n));
	return {
		...data,
		nodes,
		drawingDirty: true,
		drawingShapes: patchDiagramShape(data.drawingShapes, data.nodes, nodeId, (shape) => ({
			...shape,
			text: newText,
		})),
	};
}
