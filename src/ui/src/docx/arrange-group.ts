import {
	groupDrawingShapes,
	placeDrawingBox,
	ungroupDrawingShapes,
	type DrawingBox,
	type PicturePlacement,
	type ShapeSpec,
} from 'ooxml-core/docx';
import { Fragment, type Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { drawingGeometry } from './arrange-geometry';
import type { EditorCore } from './editor-core';
import { parsePlacement, shapeSpecOf } from './print-float-locate';

/** Makes room for the child stack at the group's place in its text layer. */
export function expandStack(
	tr: Transaction,
	pos: number,
	placement: PicturePlacement,
	count: number,
): number {
	const first = placement.relativeHeight ?? 0;
	const shift = count - 1;
	if (shift <= 0) return first;
	tr.doc.descendants((node, at) => {
		if (node.type.name !== 'image') return true;
		const anchor = parsePlacement(node.attrs.placement);
		const height = anchor?.relativeHeight ?? 0;
		if (
			node.attrs.anchored &&
			anchor &&
			Boolean(anchor.behindText) === Boolean(placement.behindText) &&
			// Equal heights paint in document order.
			(height > first || (height === first && at > pos))
		) {
			tr.setNodeAttribute(
				at,
				'placement',
				JSON.stringify({ ...anchor, relativeHeight: height + shift }),
			);
		}
		return false;
	});
	return first;
}

function shapeNode(
	node: ProseMirrorNode,
	placement: PicturePlacement,
	box: DrawingBox,
	shape: ShapeSpec,
	identity?: { id: string; name: string },
) {
	return node.type.create({
		...node.attrs,
		drawingId: identity?.id ?? null,
		drawingName: identity?.name ?? null,
		widthPx: box.width,
		heightPx: box.height,
		placement: JSON.stringify(placement),
		shape: JSON.stringify(shape),
		unsupported: 'Shape',
	});
}

export function groupSelectedShapes(
	core: EditorCore,
	view: EditorView,
	tr: Transaction,
	positions: readonly number[],
	pos: number,
	action: string,
) {
	const doc = view.state.doc;
	const current = doc.nodeAt(pos);
	if (!current || current.attrs.altText || current.attrs.title) return;
	const placement = parsePlacement(current.attrs.placement);
	const shape = shapeSpecOf(current);
	if (!placement || !shape) return;
	const geometry = drawingGeometry(core, view, pos);
	if (!geometry) return;
	if (action === 'ungroup') {
		const parts = ungroupDrawingShapes(geometry.box, shape);
		if (!parts.length) return;
		const firstHeight = expandStack(tr, pos, placement, parts.length);
		tr.replaceWith(
			pos,
			pos + current.nodeSize,
			Fragment.fromArray(
				parts.map(({ box, shape, identity }, index) =>
					shapeNode(
						current,
						{
							...placeDrawingBox(placement, box, geometry.origins),
							relativeHeight: firstHeight + index,
						},
						box,
						shape,
						identity,
					),
				),
			),
		);
		tr.setSelection(NodeSelection.create(tr.doc, pos));
		return;
	}
	if (action !== 'group' || positions.length < 2 || positions.length > 100) return;
	const parent = doc.resolve(pos).parent;
	const members = positions.flatMap((at) => {
		const node = doc.nodeAt(at);
		const spec = node && shapeSpecOf(node);
		const prior = parsePlacement(node?.attrs.placement);
		if (
			!node?.attrs.anchored ||
			!spec?.editable ||
			spec.groupChildren ||
			spec.text?.length ||
			!prior ||
			node.attrs.textBoxContent ||
			node.attrs.altText ||
			node.attrs.title ||
			node.marks.length ||
			doc.resolve(at).parent !== parent
		)
			return [];
		const located = drawingGeometry(core, view, at);
		return located?.sheet === geometry.sheet
			? [
					{
						pos: at,
						size: node.nodeSize,
						height: prior.relativeHeight ?? at,
						box: located.box,
						shape: spec,
					},
				]
			: [];
	});
	if (members.length !== positions.length) return;
	const group = groupDrawingShapes(members.sort((a, b) => a.height - b.height));
	if (!group) return;
	const origins = {
		...(group.box.x !== geometry.box.x || group.box.width !== geometry.box.width
			? { x: geometry.origins.x }
			: {}),
		...(group.box.y !== geometry.box.y || group.box.height !== geometry.box.height
			? { y: geometry.origins.y }
			: {}),
	};
	for (const member of members.sort((a, b) => b.pos - a.pos))
		tr.delete(member.pos, member.pos + member.size);
	const anchor = tr.mapping.map(pos);
	tr.insert(
		anchor,
		shapeNode(current, placeDrawingBox(placement, group.box, origins), group.box, group.shape),
	);
	tr.setSelection(NodeSelection.create(tr.doc, anchor));
}
