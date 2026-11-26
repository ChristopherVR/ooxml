import {
	alignDrawingBox,
	alignDrawingToPage,
	distributeDrawingBoxes,
	drawingBounds,
	drawingMovement,
	drawingPosition,
	drawingStackOrder,
	drawingWrapOf,
	placeDrawingBox,
	rotateDrawing,
	setDrawingWrap,
	type DrawingAlignment,
	type DrawingAlignTarget,
	type DrawingPosition,
	type DrawingTransform,
	type InlineImage,
	type PicturePlacement,
	type TextBoxContent,
} from 'ooxml-core/docx';
import { wordYjsPluginKey } from 'ooxml-core/docx/ui';
import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { drawingGeometry } from './arrange-geometry';
import { storyDrawings, toggleArrangePane } from './arrange-pane';
import { arrangementDialog } from './arrange-dialog';
import { groupSelectedShapes } from './arrange-group';
import { drawingPositions, rememberDrawingPositions } from './arrange-selection';
import { isDrawingWrap, saveDrawingWrap } from './drawing-layout-default';
import type { EditorCore } from './editor-core';
import { focusView } from './focus-view';
import { floatOrigin } from './picture-format-commands';
import { parsePlacement, shapeSpecOf } from './print-float-locate';
import { tablePropertiesContext } from './table-properties';

const targets = new WeakMap<EditorView, DrawingAlignTarget>();
const alignments: readonly DrawingAlignment[] = [
	'left',
	'center',
	'right',
	'top',
	'middle',
	'bottom',
];

function isAlignment(value: string): value is DrawingAlignment {
	return alignments.some((alignment) => alignment === value);
}

export function drawingImage(node: ProseMirrorNode | null): InlineImage | undefined {
	if (!node || node.type.name !== 'image') return undefined;
	const placement = parsePlacement(node.attrs.placement);
	return {
		relId: String(node.attrs.relId),
		partName: String(node.attrs.partName),
		contentType: String(node.attrs.contentType),
		widthPx: Number(node.attrs.widthPx),
		heightPx: Number(node.attrs.heightPx),
		anchored: Boolean(node.attrs.anchored),
		...(placement ? { placement } : {}),
	};
}

function applyPlacement(tr: Transaction, pos: number, placement: PicturePlacement | undefined) {
	tr.setNodeAttribute(pos, 'anchored', Boolean(placement));
	tr.setNodeAttribute(pos, 'placement', placement ? JSON.stringify(placement) : null);
}

function commit(view: EditorView, tr: Transaction, positions?: readonly number[]) {
	if (!tr.docChanged) return;
	const { selection } = view.state;
	if (
		selection instanceof NodeSelection &&
		tr.doc.nodeAt(selection.from)?.type.name === 'image' &&
		!tr.selectionSet
	) {
		tr.setSelection(NodeSelection.create(tr.doc, selection.from));
		if (positions && positions.length > 1)
			rememberDrawingPositions(tr.doc, selection.from, [...positions]);
	}
	const history = wordYjsPluginKey.getState(view.state);
	history?.stopCapturing();
	view.dispatch(closeHistory(tr));
	history?.stopCapturing();
}

function align(
	core: EditorCore,
	view: EditorView,
	tr: Transaction,
	positions: readonly number[],
	value: string,
) {
	if (value === 'page' || value === 'margin' || value === 'selected') {
		targets.set(view, value);
		return;
	}
	const target = targets.get(view) ?? (positions.length > 1 ? 'selected' : 'margin');
	const drawings = positions.map((pos) => ({
		pos,
		placement: drawingImage(view.state.doc.nodeAt(pos))?.placement,
	}));
	if (!drawings.every((drawing) => drawing.placement)) return;
	if (target !== 'selected' && isAlignment(value)) {
		for (const { pos, placement } of drawings)
			if (placement) applyPlacement(tr, pos, alignDrawingToPage(placement, target, value));
		return;
	}
	const items = drawings.flatMap(({ pos, placement }) => {
		const geometry = placement && drawingGeometry(core, view, pos, placement);
		return placement && geometry ? [{ pos, placement, ...geometry }] : [];
	});
	const first = items[0];
	if (
		!first ||
		items.length !== positions.length ||
		items.some((item) => first.sheet !== item.sheet)
	)
		return;
	const boxes = items.map(({ box }) => box);
	const bounds =
		target === 'page' ? first.pageBox : target === 'margin' ? first.margins : drawingBounds(boxes);
	if (!bounds) return;
	const next =
		value === 'x' || value === 'y'
			? items.length >= 3
				? distributeDrawingBoxes(boxes, bounds, value)
				: []
			: isAlignment(value) && items.length >= 2
				? boxes.map((box) => alignDrawingBox(box, bounds, value))
				: [];
	next.forEach((box, i) => {
		const item = items[i];
		if (item) applyPlacement(tr, item.pos, placeDrawingBox(item.placement, box));
	});
}

function order(
	view: EditorView,
	tr: Transaction,
	pos: number,
	value: string,
	placement: PicturePlacement,
) {
	if (value !== 'forward' && value !== 'backward' && value !== 'first' && value !== 'last') return;
	const items = storyDrawings(view).filter(
		(item) =>
			view.state.doc.nodeAt(item.pos)?.attrs.anchored &&
			item.behind === Boolean(placement.behindText),
	);
	const next = drawingStackOrder(
		items.map(({ height }) => height),
		items.findIndex((item) => item.pos === pos),
		value,
	);
	next?.forEach((height, i) => {
		const item = items[i];
		if (!item) return;
		const old = parsePlacement(view.state.doc.nodeAt(item.pos)?.attrs.placement);
		if (old && old.relativeHeight !== height)
			applyPlacement(tr, item.pos, { ...old, relativeHeight: height });
	});
}

function direction(
	view: EditorView,
	tr: Transaction,
	node: ProseMirrorNode | null,
	pos: number,
	value: string,
) {
	if (value !== 'horizontal' && value !== 'clockwise' && value !== 'counterclockwise') return;
	if (node?.attrs.textBoxContent) {
		const { vertical: _old, ...content }: TextBoxContent = JSON.parse(node.attrs.textBoxContent);
		const vertical =
			value === 'clockwise' ? 'vert' : value === 'counterclockwise' ? 'vert270' : undefined;
		tr.setNodeAttribute(
			pos,
			'textBoxContent',
			JSON.stringify({ ...content, ...(vertical ? { vertical } : {}) }),
		);
		return;
	}
	const textDirection =
		value === 'clockwise' ? 'tbRl' : value === 'counterclockwise' ? 'btLr' : 'lrTb';
	for (const cell of tablePropertiesContext(view.state)?.cells ?? [])
		tr.setNodeAttribute(cell.pos, 'textDirection', textDirection);
}

function rotate(
	view: EditorView,
	pos: number,
	node: ProseMirrorNode,
	value: 'left' | 'right' | 'flipH' | 'flipV' | number,
) {
	const shape = shapeSpecOf(node);
	const transform: DrawingTransform =
		shape ?? (node.attrs.transform ? JSON.parse(node.attrs.transform) : {});
	const next = rotateDrawing(transform, value);
	commit(
		view,
		view.state.tr.setNodeAttribute(pos, shape ? 'shape' : 'transform', JSON.stringify(next)),
	);
}

export function runArrange(
	core: EditorCore,
	view: EditorView,
	action: { key: string; value: string },
): void {
	const { key, value } = action;
	if (key === 'pane') return toggleArrangePane(core, view);
	if (!view.editable || core.readOnly) return;
	const { selection, doc } = view.state;
	const positions = drawingPositions(view.state);
	const pos = selection.from;
	const node = positions.length ? doc.nodeAt(pos) : null;
	const tr = view.state.tr;
	if (key === 'direction') {
		direction(view, tr, node, pos, value);
		commit(view, tr);
		focusView(view);
		return;
	}
	const image = drawingImage(node);
	if (!node || !image) return;
	const placement = image.placement;
	if (key === 'alt') {
		arrangementDialog(core, view, pos, 'alt', String(node.attrs.altText ?? ''), (text) =>
			commit(view, view.state.tr.setNodeAttribute(pos, 'altText', text || null)),
		);
		return;
	}
	if (key === 'rotate') {
		if (value === 'more') {
			const shape = shapeSpecOf(node);
			const transform: DrawingTransform =
				shape ?? (node.attrs.transform ? JSON.parse(node.attrs.transform) : {});
			arrangementDialog(core, view, pos, 'rotate', String(transform.rotation ?? 0), (text) =>
				rotate(view, pos, node, Number(text)),
			);
			return;
		}
		if (value === 'left' || value === 'right' || value === 'flipH' || value === 'flipV')
			rotate(view, pos, node, value);
	} else if ((key === 'wrap' || key === 'position' || key === 'order') && isDrawingWrap(value)) {
		applyPlacement(tr, pos, setDrawingWrap(image, value, floatOrigin(view, pos)).placement);
	} else if (key === 'position' && /^(top|center|bottom)-(left|center|right)$/.test(value)) {
		applyPlacement(tr, pos, drawingPosition(image, value as DrawingPosition));
	} else if (key === 'wrap' && value === 'default') {
		saveDrawingWrap(
			view.dom.ownerDocument,
			node.attrs.shape ? 'shape' : 'picture',
			drawingWrapOf(image.anchored, placement),
		);
	} else if (key === 'wrap' && placement && (value === 'move' || value === 'fix')) {
		const geometry = drawingGeometry(core, view, pos, placement);
		if (geometry)
			applyPlacement(
				tr,
				pos,
				drawingMovement(placement, value, geometry.box.y, geometry.paragraphY),
			);
	} else if (key === 'align') align(core, view, tr, positions, value);
	else if (key === 'order' && placement) order(view, tr, pos, value, placement);
	else if (key === 'group' && placement) groupSelectedShapes(core, view, tr, positions, pos, value);
	commit(view, tr, positions);
	focusView(view);
}
