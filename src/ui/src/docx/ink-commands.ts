// Word ink as ProseMirror edits: a stroke becomes (or joins) a floating ink drawing placed on its
// page, the eraser and the lasso change one drawing's strokes, and Ink to Shape puts a shape in
// their place. Every gesture is one transaction, so one undo step.
import type { InkBox, InkShape, InkStroke } from 'ooxml-core/drawingml';
import { inkTools } from '../ink/tools';
import { newShape, type DocxInk, type PicturePlacement, type ShapeSpec } from 'ooxml-core/docx';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { inkOf } from './ink-node-view';
import { nextFloatHeight } from './print-float-locate';
import { schema } from './schema';

/** Where on a page a drawing goes: the page-relative box and the paragraph that anchors it. */
export interface PagePlacement {
	blockId: string;
	/** CSS pixels from the page's top-left corner. */
	box: InkBox;
}

/** The position of the paragraph with block id `blockId` in the body, or undefined. */
export function paragraphStart(doc: ProseMirrorNode, blockId: string): number | undefined {
	let found: number | undefined;
	doc.descendants((node, pos) => {
		if (found !== undefined) return false;
		if (node.type.name === 'paragraph' && String(node.attrs.id) === blockId) {
			found = pos + 1;
			return false;
		}
		return true;
	});
	return found;
}

/** `box` on whole pixels, covering it: offsets and sizes are written in whole pixels. */
function snap(box: InkBox): InkBox {
	const x = Math.floor(box.x);
	const y = Math.floor(box.y);
	return {
		x,
		y,
		width: Math.max(1, Math.ceil(box.x + box.width - x)),
		height: Math.max(1, Math.ceil(box.y + box.height - y)),
	};
}

/** A floating placement at `box` measured from the page, in front of the text like Word's ink. */
function pagePlacement(box: InkBox, doc: ProseMirrorNode): PicturePlacement {
	return {
		wrap: 'none',
		relativeFrom: 'page',
		relativeFromV: 'page',
		offsetXPx: Math.floor(box.x),
		offsetYPx: Math.floor(box.y),
		relativeHeight: nextFloatHeight(doc),
	};
}

/**
 * Adds strokes drawn on a page (page coordinates) as a new floating ink drawing anchored to
 * `place.blockId`. Returns the position of the new node, or undefined when it cannot go there.
 */
export function insertInk(
	view: EditorView,
	strokes: readonly InkStroke[],
	place: Omit<PagePlacement, 'box'>,
): number | undefined {
	const bounds = inkTools().inkStrokesBounds(strokes);
	const at = paragraphStart(view.state.doc, place.blockId);
	if (!bounds || at === undefined || !view.editable) return undefined;
	const box = snap(bounds);
	const { width, height } = box;
	const ink: DocxInk = {
		strokes: inkTools().transformInkStrokes(strokes, -box.x, -box.y),
		width,
		height,
	};
	const node = schema.nodes.image!.create({
		unsupported: 'Ink',
		widthPx: width,
		heightPx: height,
		anchored: true,
		drawingName: 'Ink',
		placement: JSON.stringify(pagePlacement(box, view.state.doc)),
		ink: JSON.stringify(ink),
	});
	view.dispatch(closeHistory(view.state.tr.insert(at, node)));
	return at;
}

/** The page box of a floating ink drawing, when it is placed from the page. */
export function inkPageBox(node: ProseMirrorNode): InkBox | undefined {
	let placement: PicturePlacement | undefined;
	try {
		placement = JSON.parse(String(node.attrs.placement)) as PicturePlacement;
	} catch {
		return undefined;
	}
	if (placement?.relativeFrom !== 'page' || placement.relativeFromV !== 'page') return undefined;
	return {
		x: placement.offsetXPx ?? 0,
		y: placement.offsetYPx ?? 0,
		width: Number(node.attrs.widthPx),
		height: Number(node.attrs.heightPx),
	};
}

/**
 * Adds `strokes` (page coordinates) to the page-placed ink at `pos`, growing its box, as Word
 * collects strokes drawn one after another into one ink drawing. False when it is not such ink.
 */
export function appendInk(view: EditorView, pos: number, strokes: readonly InkStroke[]): boolean {
	const node = view.state.doc.nodeAt(pos);
	const ink = node && inkOf(node);
	const box = node && inkPageBox(node);
	if (!node || !ink || !box) return false;
	// The kept strokes in page coordinates (stretched as the drawing shows them), then the new ones.
	const scaled = inkTools().transformInkStrokes(
		ink.strokes,
		0,
		0,
		box.width / (ink.width || 1),
		box.height / (ink.height || 1),
	);
	const pageStrokes = [...inkTools().transformInkStrokes(scaled, box.x, box.y), ...strokes];
	const bounds = inkTools().inkStrokesBounds(pageStrokes);
	if (!bounds) return false;
	const next = snap(bounds);
	const { width, height } = next;
	const placement = JSON.parse(String(node.attrs.placement)) as PicturePlacement;
	view.dispatch(
		closeHistory(
			view.state.tr.setNodeMarkup(pos, undefined, {
				...node.attrs,
				widthPx: width,
				heightPx: height,
				placement: JSON.stringify({
					...placement,
					offsetXPx: next.x,
					offsetYPx: next.y,
				}),
				ink: JSON.stringify({
					...ink,
					strokes: inkTools().transformInkStrokes(pageStrokes, -next.x, -next.y),
					width,
					height,
				} satisfies DocxInk),
			}),
		),
	);
	return true;
}

/** New strokes for ink drawings (drawing coordinates); a drawing left with none is deleted. */
export interface InkEdit {
	pos: number;
	strokes: readonly InkStroke[];
}

function editInk(tr: Transaction, edits: readonly InkEdit[]): Transaction {
	// Last first, so a deleted drawing does not move the ones before it.
	for (const { pos, strokes } of [...edits].sort((a, b) => b.pos - a.pos)) {
		const node = tr.doc.nodeAt(pos);
		const ink = node && inkOf(node);
		if (!node || !ink) continue;
		if (strokes.length)
			tr.setNodeMarkup(pos, undefined, {
				...node.attrs,
				ink: JSON.stringify({ ...ink, strokes: [...strokes] } satisfies DocxInk),
			});
		else tr.delete(pos, pos + node.nodeSize);
	}
	return tr;
}

/**
 * Applies stroke edits as one step; `group` lets an eraser gesture's edits join one undo step
 * (they are dispatched while the pointer moves).
 */
export function applyInkEdits(view: EditorView, edits: readonly InkEdit[], group = false): void {
	const tr = editInk(view.state.tr, edits);
	if (tr.docChanged) view.dispatch(group ? tr : closeHistory(tr));
}

function shapeNode(shape: InkShape, doc: ProseMirrorNode): ProseMirrorNode {
	const spec: ShapeSpec = {
		...newShape(shape.preset),
		fill: { kind: 'none' },
		line: shape.line,
		...(shape.rotation ? { rotation: shape.rotation } : {}),
		...(shape.flipH ? { flipH: true } : {}),
		...(shape.flipV ? { flipV: true } : {}),
	};
	return schema.nodes.image!.create({
		unsupported: 'Shape',
		widthPx: Math.max(1, Math.round(shape.box.width)),
		heightPx: Math.max(1, Math.round(shape.box.height)),
		anchored: true,
		placement: JSON.stringify(pagePlacement(shape.box, doc)),
		shape: JSON.stringify(spec),
	});
}

/**
 * Ink to Shape as one step: the recognised strokes leave their drawings (`edits`) and each shape
 * (page coordinates) goes on the page, anchored to its paragraph.
 */
export function convertInkToShapes(
	view: EditorView,
	edits: readonly InkEdit[],
	shapes: readonly { shape: InkShape; blockId: string }[],
): boolean {
	if (!view.editable || !shapes.length) return false;
	const tr = editInk(view.state.tr, edits);
	for (const { shape, blockId } of shapes) {
		const at = paragraphStart(tr.doc, blockId);
		if (at !== undefined) tr.insert(at, shapeNode(shape, tr.doc));
	}
	view.dispatch(closeHistory(tr));
	return true;
}
