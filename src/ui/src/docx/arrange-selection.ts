import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import { isPictureNode } from './picture-format-commands';

export function isArrangeDrawing(
	node: ProseMirrorNode | null | undefined,
): node is ProseMirrorNode {
	return Boolean(node?.type.name === 'image' && (node.attrs.shape || isPictureNode(node)));
}

// Each selection is valid only for the exact version of the story it came from.
const selections = new WeakMap<ProseMirrorNode, { primary: number; positions: number[] }>();

export function drawingPositions(state: EditorState): readonly number[] {
	const { selection, doc } = state;
	if (!(selection instanceof NodeSelection) || !isArrangeDrawing(selection.node)) return [];
	const stored = selections.get(doc);
	return stored?.primary === selection.from ? stored.positions : [selection.from];
}

export function rememberDrawingPositions(
	doc: ProseMirrorNode,
	primary: number,
	positions: number[],
): void {
	selections.set(doc, { primary, positions });
}
