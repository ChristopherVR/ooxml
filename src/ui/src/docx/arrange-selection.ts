import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, type EditorState, type Selection } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { isPictureNode } from './picture-format-commands';

export function isArrangeDrawing(
	node: ProseMirrorNode | null | undefined,
): node is ProseMirrorNode {
	return Boolean(node?.type.name === 'image' && (node.attrs.shape || isPictureNode(node)));
}

/** A node selection with additional drawings, remapped by the same transactions as its primary. */
class DrawingSelection extends NodeSelection {
	constructor(
		doc: ProseMirrorNode,
		readonly positions: readonly number[],
		primary: number,
	) {
		super(doc.resolve(primary));
	}

	override eq(other: Selection): boolean {
		return (
			other instanceof DrawingSelection &&
			this.anchor === other.anchor &&
			this.positions.length === other.positions.length &&
			this.positions.every((pos, index) => pos === other.positions[index])
		);
	}

	override map(doc: ProseMirrorNode, mapping: Parameters<NodeSelection['map']>[1]): Selection {
		const positions = this.positions.flatMap((pos) => {
			const mapped = mapping.mapResult(pos);
			return !mapped.deleted && isArrangeDrawing(doc.nodeAt(mapped.pos)) ? [mapped.pos] : [];
		});
		const primary = mapping.map(this.anchor);
		const fallback = positions[0];
		if (fallback === undefined) return super.map(doc, mapping);
		const anchor = positions.includes(primary) ? primary : fallback;
		return positions.length === 1
			? NodeSelection.create(doc, anchor)
			: new DrawingSelection(doc, positions, anchor);
	}
}

export function drawingPositions(state: EditorState): readonly number[] {
	const { selection } = state;
	if (selection instanceof DrawingSelection) return selection.positions;
	return selection instanceof NodeSelection && isArrangeDrawing(selection.node)
		? [selection.from]
		: [];
}

export function selectDrawing(view: EditorView, pos: number, extend: boolean): void {
	const { doc } = view.state;
	if (!isArrangeDrawing(doc.nodeAt(pos))) return;
	const previous = drawingPositions(view.state);
	const next = extend
		? previous.includes(pos)
			? previous.filter((item) => item !== pos)
			: [...previous, pos]
		: [pos];
	const primary = next.includes(pos) ? pos : (next[0] ?? pos);
	const selection =
		next.length > 1 ? new DrawingSelection(doc, next, primary) : NodeSelection.create(doc, primary);
	view.dispatch(view.state.tr.setSelection(selection));
}

export function extendsDrawingSelection(event: MouseEvent): boolean {
	return event.ctrlKey || event.metaKey || event.shiftKey;
}

export function selectOnDrawingMouseDown(view: EditorView, event: MouseEvent): boolean {
	if (
		event.button !== 0 ||
		!extendsDrawingSelection(event) ||
		view.dom.classList.contains('dve-drawing-shape') ||
		!(event.target instanceof Element)
	)
		return false;
	const drawing = event.target.closest('.dve-shape, .dve-picture');
	if (!drawing || !view.dom.contains(drawing)) return false;
	const pos = view.posAtDOM(drawing, 0);
	const node = view.state.doc.nodeAt(pos);
	if (!isArrangeDrawing(node) || !node.attrs.anchored) return false;
	// Handle the press before ProseMirror replaces the existing selection.
	event.preventDefault();
	selectDrawing(view, pos, true);
	view.focus();
	return true;
}

export function secondaryDrawingDecorations(state: EditorState): DecorationSet {
	if (!(state.selection instanceof DrawingSelection)) return DecorationSet.empty;
	return DecorationSet.create(
		state.doc,
		state.selection.positions.flatMap((pos) => {
			const node = state.doc.nodeAt(pos);
			return node && pos !== state.selection.from
				? [Decoration.node(pos, pos + node.nodeSize, { class: 'dve-drawing-secondary' })]
				: [];
		}),
	);
}
