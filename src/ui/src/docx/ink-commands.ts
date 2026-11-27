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
import { drawHelpers } from './draw-helpers';
import { schema } from './schema';

const inkOf = (node: ProseMirrorNode) => drawHelpers().inkOf(node);

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
		relativeHeight: drawHelpers().nextFloatHeight(doc),
	};
}

/**
 * Adds strokes drawn on a page (page coordinates) as a new floating ink drawing anchored to
 * `place.blockId`. Returns the position of the new node, or undefined when it cannot go there.
 */
export function insertInk(
	view: EditorView,
	strokes: readonly InkStroke[],
	place: Omit<PagePlacement, 'box'> & { name?: string },
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
		// Its name, saved as `wp:docPr/@name` and read by assistive technology, in the editor's language.
		drawingName: place.name ?? 'Ink',
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

/**
 * The attributes of ink drawing `node` holding `strokes` (its own coordinates), its box fitted
 * to them as appendInk fits a growing one: erasing the edge of a drawing shrinks its box and
 * moves its offsets, so it never keeps a frame larger than its ink. Offsets set by an alignment
 * stay as they are.
 */
function fitted(node: ProseMirrorNode, ink: DocxInk, strokes: readonly InkStroke[]) {
	const tools = inkTools();
	const bounds = tools.inkStrokesBounds(strokes);
	const width = Number(node.attrs.widthPx) || 1;
	const height = Number(node.attrs.heightPx) || 1;
	const sx = width / (ink.width || 1);
	const sy = height / (ink.height || 1);
	const unchanged = { ...node.attrs, ink: JSON.stringify({ ...ink, strokes: [...strokes] }) };
	if (!bounds) return unchanged;
	const box = snap({
		x: bounds.x * sx,
		y: bounds.y * sy,
		width: bounds.width * sx,
		height: bounds.height * sy,
	});
	if (!box.x && !box.y && box.width === width && box.height === height) return unchanged;
	const placement = inkPlacement(node);
	return {
		...node.attrs,
		widthPx: box.width,
		heightPx: box.height,
		...(node.attrs.anchored && placement
			? {
					placement: JSON.stringify({
						...placement,
						...(placement.align ? {} : { offsetXPx: (placement.offsetXPx ?? 0) + box.x }),
						...(placement.alignV ? {} : { offsetYPx: (placement.offsetYPx ?? 0) + box.y }),
					}),
				}
			: {}),
		ink: JSON.stringify({
			...ink,
			strokes: tools.transformInkStrokes(strokes, -box.x / sx, -box.y / sy),
			width: box.width / sx,
			height: box.height / sy,
		} satisfies DocxInk),
	};
}

function inkPlacement(node: ProseMirrorNode): PicturePlacement | undefined {
	try {
		return JSON.parse(String(node.attrs.placement)) as PicturePlacement;
	} catch {
		return undefined;
	}
}

function editInk(tr: Transaction, edits: readonly InkEdit[], fit: boolean): Transaction {
	// Last first, so a deleted drawing does not move the ones before it.
	for (const { pos, strokes } of [...edits].sort((a, b) => b.pos - a.pos)) {
		const node = tr.doc.nodeAt(pos);
		const ink = node && inkOf(node);
		if (!node || !ink) continue;
		if (!strokes.length) tr.delete(pos, pos + node.nodeSize);
		else
			tr.setNodeMarkup(
				pos,
				undefined,
				fit
					? fitted(node, ink, strokes)
					: { ...node.attrs, ink: JSON.stringify({ ...ink, strokes: [...strokes] }) },
			);
	}
	return tr;
}

/**
 * Applies stroke edits as one step, each drawing's box fitted to what it keeps. An eraser gesture
 * passes `group` after its first edit, so the gesture is one undo step, and `fit: false` while the
 * pointer moves (the pages still show the boxes it is erasing against); `fitInk` fits them at the
 * end. Returns the transaction, to map positions through.
 */
export function applyInkEdits(
	view: EditorView,
	edits: readonly InkEdit[],
	options: { group?: boolean; fit?: boolean } = {},
): Transaction | undefined {
	const tr = editInk(view.state.tr, edits, options.fit ?? true);
	if (!tr.docChanged) return undefined;
	view.dispatch(options.group ? tr : closeHistory(tr));
	return tr;
}

/**
 * Fits the boxes of the ink drawings at `positions` to their strokes, in the undo step of the
 * edit dispatched at `time` (the eraser gesture's last), however long the pointer rested since.
 */
export function fitInk(view: EditorView, positions: Iterable<number>, time: number): void {
	const tr = view.state.tr.setTime(time);
	for (const pos of positions) {
		const node = tr.doc.nodeAt(pos);
		const ink = node && inkOf(node);
		if (node && ink) tr.setNodeMarkup(pos, undefined, fitted(node, ink, ink.strokes));
	}
	if (tr.docChanged) view.dispatch(tr);
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
	const tr = editInk(view.state.tr, edits, true);
	for (const { shape, blockId } of shapes) {
		const at = paragraphStart(tr.doc, blockId);
		if (at !== undefined) tr.insert(at, shapeNode(shape, tr.doc));
	}
	view.dispatch(closeHistory(tr));
	return true;
}
