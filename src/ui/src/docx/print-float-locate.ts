import type { PicturePlacement, ShapeSpec } from 'ooxml-core/docx';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { FloatBox, FloatHandle } from './print-float-geometry';

export interface Located {
	pos: number;
	node: ProseMirrorNode;
	blockId: string;
	index: number;
}

/** A float as the layout counts them (`paragraphFloats`): anchored with media bytes or a shape. */
const countsAsFloat = (node: ProseMirrorNode): boolean =>
	node.type.name === 'image' &&
	Boolean(node.attrs.anchored && (node.attrs.partName || node.attrs.shape));

/** A float the page can select: an anchored picture with media bytes, or an anchored shape. */
export const isFloat = (node: ProseMirrorNode): boolean => countsAsFloat(node);

/** Whether a float is a drawing shape or text box rather than a picture. */
export const isShapeFloat = (node: ProseMirrorNode): boolean =>
	isFloat(node) && Boolean(node.attrs.shape);

/** The shape spec of a float, when it is a shape. */
export function shapeSpecOf(node: ProseMirrorNode): ShapeSpec | undefined {
	if (typeof node.attrs.shape !== 'string') return undefined;
	try {
		return JSON.parse(node.attrs.shape) as ShapeSpec;
	} catch {
		return undefined;
	}
}

/** The float `index` of a paragraph, in run order. */
export function locateFloat(
	doc: ProseMirrorNode,
	blockId: string,
	index: number,
): Located | undefined {
	let found: Located | undefined;
	doc.descendants((node, pos) => {
		if (found) return false;
		if (node.type.name !== 'paragraph') return true;
		if (String(node.attrs.id) !== blockId) return false;
		let count = 0;
		node.forEach((child, offset) => {
			if (!countsAsFloat(child)) return;
			if (count++ === index && isFloat(child))
				found = { pos: pos + 1 + offset, node: child, blockId, index };
		});
		return false;
	});
	return found;
}

/** The floating object a node selection holds, with its paragraph id and index. */
export function selectedFloat(state: EditorState): Located | undefined {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || !isFloat(selection.node)) return undefined;
	const $pos = state.doc.resolve(selection.from);
	const paragraph = $pos.parent;
	if (paragraph.type.name !== 'paragraph') return undefined;
	let index = 0;
	paragraph.forEach((child, offset) => {
		if ($pos.start() + offset < selection.from && countsAsFloat(child)) index++;
	});
	return {
		pos: selection.from,
		node: selection.node,
		blockId: String(paragraph.attrs.id),
		index,
	};
}

export const number = (value: string | undefined): number => Number.parseFloat(value ?? '') || 0;

export function boxOf(element: HTMLElement): FloatBox {
	return {
		x: number(element.style.left),
		y: number(element.style.top),
		width: number(element.style.width),
		height: number(element.style.height),
	};
}

export function place(element: HTMLElement, box: FloatBox): void {
	Object.assign(element.style, {
		left: `${box.x}px`,
		top: `${box.y}px`,
		width: `${box.width}px`,
		height: `${box.height}px`,
	});
}

export const CURSORS: Record<FloatHandle, string> = {
	nw: 'nwse-resize',
	se: 'nwse-resize',
	ne: 'nesw-resize',
	sw: 'nesw-resize',
	n: 'ns-resize',
	s: 'ns-resize',
	e: 'ew-resize',
	w: 'ew-resize',
};

export function parsePlacement(value: unknown): PicturePlacement | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as PicturePlacement;
	} catch {
		return undefined;
	}
}
