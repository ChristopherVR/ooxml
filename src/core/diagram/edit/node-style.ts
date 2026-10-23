/**
 * SmartArt per-node visual-override editing operation.
 *
 * Like the other node ops, the function here is pure and immutable: it returns
 * a new data object of the caller's type and patches the matching drawing
 * shape (or marks the shapes stale) so the renderer honours the per-node fill /
 * font override. Moved from `pptx/core/utils/smartart-editing-node-style.ts`
 * and made generic.
 *
 * @module diagram/edit/node-style
 */

import type { DiagramNode, DiagramNodeStyle } from '../model/node';
import { findDiagramShapeIndex } from './shape-match';
import type { DiagramEditShape, DiagramEditableData } from './types';

/** A style patch: present keys overwrite, keys set to `undefined` are removed. */
export type DiagramNodeStylePatch = {
	[K in keyof DiagramNodeStyle]?: DiagramNodeStyle[K] | undefined;
};

/** The drawing-shape fields a node style override patches. */
export interface DiagramStyledShape extends DiagramEditShape {
	fillColor?: string | undefined;
	strokeColor?: string | undefined;
	fontColor?: string | undefined;
}

/** Merge a partial style onto an existing one, dropping cleared (undefined) keys. */
function mergeNodeStyle(
	current: DiagramNodeStyle | undefined,
	patch: DiagramNodeStylePatch,
): DiagramNodeStyle {
	const merged: DiagramNodeStyle = { ...current };
	for (const key of Object.keys(patch) as (keyof DiagramNodeStyle)[]) {
		const value = patch[key];
		if (value === undefined) {
			delete merged[key];
		} else {
			// Assign through a typed bridge so each optional field keeps its type.
			(merged as Record<keyof DiagramNodeStyle, unknown>)[key] = value;
		}
	}
	return merged;
}

/**
 * Set (merge) a per-node visual override on a SmartArt node by id.
 *
 * The `style` partial is shallow-merged onto the node's existing
 * {@link DiagramNodeStyle}: present keys overwrite, keys explicitly set to
 * `undefined` are removed. When the resulting style has no keys, the `style`
 * field is dropped from the node entirely. `drawingShapes` is cleared so the
 * renderer reflows and the new colours take effect.
 *
 * Returns a new data object; the input is never mutated. When `nodeId`
 * matches no node, the original `data` reference is returned unchanged.
 *
 * @example
 * ```ts
 * const next = setDiagramNodeStyle(data, "node-2", { fillColor: "#FF0000", bold: true });
 * ```
 */
export function setDiagramNodeStyle<D extends DiagramEditableData<DiagramNode, DiagramStyledShape>>(
	data: D,
	nodeId: string,
	style: DiagramNodeStylePatch,
): D {
	let matched = false;
	const nodes = data.nodes.map((node) => {
		if (node.id !== nodeId) {
			return node;
		}
		matched = true;
		const nextStyle = mergeNodeStyle(node.style, style);
		const hasStyle = Object.keys(nextStyle).length > 0;
		const updated = { ...node };
		if (hasStyle) {
			updated.style = nextStyle;
		} else {
			delete updated.style;
		}
		return updated;
	});

	if (!matched) {
		return data;
	}

	return {
		...data,
		nodes,
		drawingDirty: true,
		drawingShapes: patchDrawingShapeStyle(data.drawingShapes, data.nodes, nodeId, style),
	};
}

/**
 * Patch a single drawing shape's visual properties in-place (immutable copy).
 * Returns `undefined` unchanged when no drawing shapes exist.
 */
function patchDrawingShapeStyle<S extends DiagramStyledShape>(
	shapes: S[] | undefined,
	nodes: readonly DiagramNode[],
	nodeId: string,
	style: DiagramNodeStylePatch,
): S[] | undefined {
	if (!shapes || shapes.length === 0) {
		return shapes;
	}
	const idx = findDiagramShapeIndex(shapes, nodes, nodeId);
	if (idx < 0) {
		// Cannot resolve which shape corresponds to this node (e.g. PowerPoint-
		// generated shapes with opaque IDs). Mark as stale so a full rebuild
		// picks up the new style.
		return [];
	}
	const patch: Partial<DiagramStyledShape> = {};
	if (style.fillColor !== undefined) {
		patch.fillColor = style.fillColor;
	}
	if (style.lineColor !== undefined) {
		patch.strokeColor = style.lineColor;
	}
	if (style.fontColor !== undefined) {
		patch.fontColor = style.fontColor;
	}
	if (Object.keys(patch).length === 0) {
		return shapes;
	}
	const updated = [...shapes];
	updated[idx] = { ...updated[idx]!, ...patch };
	return updated;
}
