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

function shapeNode(
	node: ProseMirrorNode,
	placement: PicturePlacement,
	box: DrawingBox,
	shape: ShapeSpec,
) {
	return node.type.create({
		...node.attrs,
		widthPx: box.width,
		heightPx: box.height,
		placement: JSON.stringify(placeDrawingBox(placement, box)),
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
	const geometry = drawingGeometry(core, view, pos, placement);
	if (!geometry) return;
	if (action === 'ungroup') {
		const parts = ungroupDrawingShapes(geometry.box, shape);
		if (!parts.length) return;
		tr.replaceWith(
			pos,
			pos + current.nodeSize,
			Fragment.fromArray(parts.map(({ box, shape }) => shapeNode(current, placement, box, shape))),
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
		const located = drawingGeometry(core, view, at, prior);
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
	const first = Math.min(...positions);
	for (const member of members.sort((a, b) => b.pos - a.pos))
		tr.delete(member.pos, member.pos + member.size);
	tr.insert(first, shapeNode(current, placement, group.box, group.shape));
	tr.setSelection(NodeSelection.create(tr.doc, first));
}
