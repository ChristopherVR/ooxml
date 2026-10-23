// SmartArt in the editor: finding the selected diagram, applying a SmartArt Design or text pane
// command to it through the core (`editDocxDiagram`), and inserting a new one. Every change is one
// ProseMirror transaction on the image node's `diagram` attribute, so undo and redo go through the
// editor's history and collaboration like any other edit.
import { closeHistory } from 'prosemirror-history';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	editDocxDiagram,
	type DocxDiagram,
	type DocxDiagramEdit,
	type InlineImage,
} from 'ooxml-core/docx';
import type { OfficeSmartArtTextNode } from '../smartart';
import { schema } from './schema';
import { parseDiagram } from './smartart-node-view';

export interface SelectedDiagram {
	/** Position of the image node that holds the diagram. */
	pos: number;
	diagram: DocxDiagram;
	widthPx: number;
	heightPx: number;
}

/** The SmartArt the selection is on (an image atom carrying a `diagram`), if any. */
export function selectedDiagram(state: EditorState): SelectedDiagram | undefined {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || selection.node.type !== schema.nodes.image)
		return undefined;
	const diagram = parseDiagram(selection.node.attrs.diagram);
	if (!diagram) return undefined;
	return {
		pos: selection.from,
		diagram,
		widthPx: Number(selection.node.attrs.widthPx),
		heightPx: Number(selection.node.attrs.heightPx),
	};
}

/** Replaces the diagram at `pos` in one undoable step and keeps it selected. */
export function replaceDiagram(view: EditorView, pos: number, diagram: DocxDiagram): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!view.editable || !node || node.type !== schema.nodes.image) return false;
	const tr = view.state.tr.setNodeMarkup(pos, undefined, {
		...node.attrs,
		diagram: JSON.stringify(diagram),
	});
	view.dispatch(closeHistory(tr.setSelection(NodeSelection.create(tr.doc, pos))));
	return true;
}

/**
 * Applies one command to the selected diagram. Returns the node the command created or acted on
 * (for the text pane's focus), or undefined when nothing changed.
 */
export function runDiagramEdit(view: EditorView, edit: DocxDiagramEdit): string | undefined {
	const selected = selectedDiagram(view.state);
	if (!selected || !view.editable) return undefined;
	const result = editDocxDiagram(selected.diagram, edit);
	if (!result.changed) return undefined;
	replaceDiagram(view, selected.pos, result.diagram);
	return result.nodeId ?? '';
}

/** Typing in the text pane: consecutive keystrokes join one undo step, as in Word. */
export function setDiagramNodeText(view: EditorView, nodeId: string, text: string): boolean {
	const selected = selectedDiagram(view.state);
	if (!selected || !view.editable) return false;
	const result = editDocxDiagram(selected.diagram, { kind: 'text', nodeId, text });
	if (!result.changed) return false;
	const node = view.state.doc.nodeAt(selected.pos);
	if (!node) return false;
	const tr = view.state.tr.setNodeMarkup(selected.pos, undefined, {
		...node.attrs,
		diagram: JSON.stringify(result.diagram),
	});
	view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, selected.pos)));
	return true;
}

/** Inserts a new SmartArt at the selection, as its own paragraph content, and selects it. */
export function insertDiagram(view: EditorView, image: InlineImage): boolean {
	if (!view.editable || !image.diagram) return false;
	const node = schema.nodes.image.create({
		relId: '',
		partName: '',
		contentType: image.contentType,
		widthPx: image.widthPx,
		heightPx: image.heightPx,
		anchored: false,
		unsupported: 'SmartArt',
		diagram: JSON.stringify(image.diagram),
	});
	const tr = view.state.tr.replaceSelectionWith(node, false);
	const pos = tr.mapping.map(view.state.selection.from, -1);
	const at = tr.doc.nodeAt(pos)?.type === schema.nodes.image ? pos : pos - 1;
	if (tr.doc.nodeAt(at)?.type === schema.nodes.image)
		tr.setSelection(NodeSelection.create(tr.doc, at));
	view.dispatch(closeHistory(tr.scrollIntoView()));
	return true;
}

/** The diagram's nodes as text pane lines, with their depth below the top level. */
export function diagramOutline(diagram: DocxDiagram): OfficeSmartArtTextNode[] {
	const ids = new Set(diagram.nodes.map((node) => node.id));
	const levels = new Map<string, number>();
	return diagram.nodes.map((node) => {
		const parent = node.parentId && ids.has(node.parentId) ? node.parentId : undefined;
		const level = parent === undefined ? 0 : (levels.get(parent) ?? 0) + 1;
		levels.set(node.id, level);
		return { id: node.id, text: node.text, level };
	});
}
