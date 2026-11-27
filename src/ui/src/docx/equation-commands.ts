// Insert > Equation: new equations from LaTeX or the gallery, and editing the selected one.
import { closeHistory } from 'prosemirror-history';
import { NodeSelection, TextSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { convertOmmlToLatex, parseOmml } from 'ooxml-core/math';
import { schema } from './schema';
import { equationMarkup } from '../math/equation-fields';
export { equationMarkup } from '../math/equation-fields';

export interface SelectedEquation {
	pos: number;
	latex: string;
	display: boolean;
}

/** The equation node the selection is on, with its content as LaTeX. */
export function selectedEquation(state: EditorState): SelectedEquation | undefined {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || selection.node.type !== schema.nodes.equation)
		return undefined;
	let latex = '';
	try {
		latex = convertOmmlToLatex(parseOmml(String(selection.node.attrs.omml)));
	} catch {
		latex = '';
	}
	return { pos: selection.from, latex, display: Boolean(selection.node.attrs.display) };
}

export type EquationResult = 'done' | 'invalid' | 'unavailable';

/** Word makes an equation typed on its own line a display equation and one in text inline. */
export function aloneOnItsLine(state: EditorState): boolean {
	const { $from } = state.selection;
	return (
		$from.parent.type.name === 'paragraph' &&
		$from.parent.textContent.trim() === '' &&
		$from.parent.childCount <= 1
	);
}

export function insertEquation(
	view: EditorView,
	latex: string,
	display = aloneOnItsLine(view.state),
): EquationResult {
	if (!view.editable) return 'unavailable';
	const omml = equationMarkup(latex, display);
	if (!omml) return 'invalid';
	const node = schema.nodes.equation!.create({ omml, display });
	view.dispatch(closeHistory(view.state.tr.replaceSelectionWith(node, false)).scrollIntoView());
	return 'done';
}

/** Replaces the equation at `pos` with new LaTeX, as one undo step, keeping it selected. */
export function updateEquation(
	view: EditorView,
	pos: number,
	latex: string,
	display: boolean,
): EquationResult {
	const node = view.state.doc.nodeAt(pos);
	if (!view.editable || node?.type !== schema.nodes.equation) return 'unavailable';
	const omml = equationMarkup(latex, display);
	if (!omml) return 'invalid';
	const tr = view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, omml, display });
	view.dispatch(closeHistory(tr.setSelection(NodeSelection.create(tr.doc, pos))));
	return 'done';
}

/**
 * Inserts an equation from its markup at the selection and selects it, so the equation editor
 * opens on it (Insert New Equation, Alt+=). Returns its position.
 */
export function insertEquationXml(
	view: EditorView,
	omml: string,
	display = aloneOnItsLine(view.state),
): number | undefined {
	if (!view.editable) return undefined;
	const node = schema.nodes.equation!.create({ omml, display });
	const tr = closeHistory(view.state.tr.replaceSelectionWith(node, false));
	const pos = tr.selection.from - node.nodeSize;
	if (tr.doc.nodeAt(pos)?.type !== schema.nodes.equation) return undefined;
	view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)).scrollIntoView());
	return pos;
}

/**
 * Ink to Math and Ink Equation: a new equation from recognized LaTeX at the caret, or after the
 * selected equation (which is never replaced), selected so its editor opens on it.
 */
export function insertEquationAfterSelection(view: EditorView, latex: string): boolean {
	if (!view.editable) return false;
	const { selection } = view.state;
	if (selection instanceof NodeSelection && selection.node.type === schema.nodes.equation)
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, selection.to)));
	const display = aloneOnItsLine(view.state);
	const omml = equationMarkup(latex, display);
	return Boolean(omml && insertEquationXml(view, omml, display) !== undefined);
}

/** Sets the markup of the equation at `pos`, keeping it selected; false when none is there. */
export function setEquationXml(view: EditorView, pos: number, omml: string): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!view.editable || node?.type !== schema.nodes.equation) return false;
	if (node.attrs.omml === omml) return true;
	const tr = view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, omml });
	view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
	return true;
}

/** Every equation in the document, as `[position, markup, display]`. */
export function equationsIn(state: EditorState): Array<[number, string, boolean]> {
	const found: Array<[number, string, boolean]> = [];
	state.doc.descendants((node, pos) => {
		if (node.type === schema.nodes.equation)
			found.push([pos, String(node.attrs.omml), Boolean(node.attrs.display)]);
	});
	return found;
}
