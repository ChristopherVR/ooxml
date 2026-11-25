// Insert > Equation: new equations from LaTeX or the gallery, and editing the selected one.
import { closeHistory } from 'prosemirror-history';
import { NodeSelection, type EditorState } from 'prosemirror-state';
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
