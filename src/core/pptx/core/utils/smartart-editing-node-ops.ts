/**
 * SmartArt node mutation operations, bound to `PptxSmartArtData`.
 *
 * The operations live in `diagram/edit` (`node-ops`, `node-order`) over the
 * neutral data model; this facade keeps the historical pptx names and mints
 * new point ids through the pptx GUID generator.
 *
 * @module smartart-editing-node-ops
 */

import {
	addDiagramNodeAbove,
	addDiagramNodeAfter,
	addDiagramNodeBefore,
	addDiagramNodeBelow,
	demoteDiagramNode,
	moveDiagramNode,
	moveDiagramNodeToIndex,
	promoteDiagramNode,
	removeDiagramNode,
	setDiagramNodeText,
	type DiagramEditOptions,
} from '../../../diagram/edit/index';
import type { PptxSmartArtData } from '../types';
import { generateFontGuid } from './font-deobfuscation';

/**
 * New `dgm:pt`/`dgm:cxn` ids must use the `{GUID}` form; PowerPoint's loader
 * rejects a file with differently formatted ids as corrupt on open.
 */
const PPTX_EDIT_OPTIONS: DiagramEditOptions = { createModelId: () => `{${generateFontGuid()}}` };

/** No-op kept for API compatibility; id generation no longer needs a counter. */
export function resetSmartArtEditCounter(): void {}

/** Add a new node after a given sibling, or at the end. */
export function addSmartArtNode(
	data: PptxSmartArtData,
	text: string,
	afterNodeId?: string,
): PptxSmartArtData {
	return addDiagramNodeAfter(data, text, afterNodeId, PPTX_EDIT_OPTIONS);
}

/** Add a new node before a given sibling (Add Shape Before), or at the end. */
export function addSmartArtNodeBefore(
	data: PptxSmartArtData,
	text: string,
	beforeNodeId?: string,
): PptxSmartArtData {
	return addDiagramNodeBefore(data, text, beforeNodeId, PPTX_EDIT_OPTIONS);
}

/**
 * Add a new node above a given node (Add Shape Above): the new node takes the
 * target's place and the target becomes its child. Unknown ids return `data`.
 */
export function addSmartArtNodeAbove(
	data: PptxSmartArtData,
	nodeId: string,
	text: string,
): PptxSmartArtData {
	return addDiagramNodeAbove(data, nodeId, text, PPTX_EDIT_OPTIONS);
}

/** Remove a node by ID, re-parenting its children. */
export function removeSmartArtNode(data: PptxSmartArtData, nodeId: string): PptxSmartArtData {
	return removeDiagramNode(data, nodeId);
}

/** Update the text of a node by ID (and its drawing shape, when one matches). */
export function updateSmartArtNodeText(
	data: PptxSmartArtData,
	nodeId: string,
	newText: string,
): PptxSmartArtData {
	return setDiagramNodeText(data, nodeId, newText);
}

/** Move a node within its sibling group (1 = down/right, -1 = up/left). */
export function reorderSmartArtNode(
	data: PptxSmartArtData,
	nodeId: string,
	direction: 1 | -1,
): PptxSmartArtData {
	return moveDiagramNode(data, nodeId, direction);
}

/** Promote a child node to be a sibling of its parent. */
export function promoteSmartArtNode(data: PptxSmartArtData, nodeId: string): PptxSmartArtData {
	return promoteDiagramNode(data, nodeId);
}

/** Demote a node to become a child of its preceding sibling. */
export function demoteSmartArtNode(data: PptxSmartArtData, nodeId: string): PptxSmartArtData {
	return demoteDiagramNode(data, nodeId);
}

/** Add a new node as a child of `parentId`, or as a root-level item. */
export function addSmartArtNodeAsChild(
	data: PptxSmartArtData,
	parentId?: string,
	text?: string,
): PptxSmartArtData {
	return addDiagramNodeBelow(data, parentId, text, PPTX_EDIT_OPTIONS);
}

/** Move a node to a specific index within its sibling group. */
export function reorderSmartArtNodeToIndex(
	data: PptxSmartArtData,
	nodeId: string,
	newIndex: number,
): PptxSmartArtData {
	return moveDiagramNodeToIndex(data, nodeId, newIndex);
}
