// Equation commands outside the editor box: structure and built-in equation text, Insert New
// Equation (Alt+=), the Equation gallery and Convert > All.
import { NodeSelection, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { EQUATION_GALLERY, type EquationInputFormat } from 'ooxml-core/math';
import { equationKit as kit } from './equation-kit';
import { groupPlaceholders } from './equation-linear-text';
import {
	aloneOnItsLine,
	equationsIn,
	insertEquationXml,
	setEquationXml,
} from './equation-commands';

/** Linear text for a structure or built-in equation (LaTeX with `⬚`) in the input format. */
export function structureText(latex: string, format: EquationInputFormat): string {
	if (format === 'latex') return latex;
	const omml = kit().buildUpEquation(latex, 'latex', false);
	const text = omml && kit().linearizeEquation(omml, 'unicode');
	return text ? groupPlaceholders(text) : latex;
}

/** The linear text of structure `group/id`, if there is one. */
export function structureById(value: string, format: EquationInputFormat): string | undefined {
	const [group = '', id = ''] = value.split('/');
	const structure = kit().equationStructure(group, id);
	return structure && structureText(structure.latex, format);
}

/** A built-in equation's markup, built up for display or inline. */
export function builtInXml(id: string, display: boolean): string | undefined {
	const entry = EQUATION_GALLERY.find((item) => item.id === id);
	return entry && kit().buildUpEquation(entry.latex, 'latex', display);
}

/**
 * The Equation gallery: the built-in equation fills the empty equation at `emptyAt`, or is
 * inserted as a new equation after the selected one (or at the caret), and selected.
 */
export function placeBuiltIn(view: EditorView, id: string, emptyAt: number | undefined): void {
	const node = emptyAt === undefined ? null : view.state.doc.nodeAt(emptyAt);
	if (node && emptyAt !== undefined) {
		const omml = builtInXml(id, Boolean(node.attrs.display));
		if (omml) setEquationXml(view, emptyAt, omml);
		return;
	}
	const { selection } = view.state;
	if (selection instanceof NodeSelection && selection.node.type.name === 'equation')
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, selection.to)));
	const alone = aloneOnItsLine(view.state);
	const omml = builtInXml(id, alone);
	if (omml) insertEquationXml(view, omml, alone);
}

/** Inserts an empty equation at the selection and selects it (Alt+=). */
export function insertNewEquation(view: EditorView): number | undefined {
	const alone = aloneOnItsLine(view.state);
	return insertEquationXml(view, kit().emptyEquationXml(alone), alone);
}

/**
 * Convert > All: every equation to Professional or to Linear, as one step. An equation whose
 * linear text cannot hold it exactly is left as it is; in a display paragraph with several
 * equations, each one is converted on its own and the others are left as they are.
 */
export function convertAllEquations(view: EditorView, linear: boolean): void {
	if (!view.editable) return;
	let tr = view.state.tr;
	for (const [pos, omml, display] of equationsIn(view.state)) {
		const next = convertedEquation(omml, display, linear);
		const node = tr.doc.nodeAt(pos);
		if (next && node && next !== omml)
			tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, omml: next });
	}
	if (!tr.docChanged) return;
	// The selected equation stays selected, so its editor and the caret stay in it.
	const { selection } = view.state;
	if (selection instanceof NodeSelection && selection.node.type.name === 'equation')
		tr = tr.setSelection(NodeSelection.create(tr.doc, tr.mapping.map(selection.from)));
	view.dispatch(tr);
}

/** An equation's markup shown in Linear or Professional; undefined when nothing changes. */
function convertedEquation(omml: string, display: boolean, linear: boolean): string | undefined {
	const parts = kit().equationParts(omml);
	if (parts < 2) {
		const edit = kit().editableEquation(omml, 'unicode');
		if (!edit || edit.linear === linear) return undefined;
		return kit().writeEditedEquation({ ...edit, format: 'unicode', display, linear });
	}
	let next = omml;
	for (let part = 0; part < parts; part++) {
		const edit = kit().editableEquation(next, 'unicode', part);
		if (!edit || edit.linear === linear) continue;
		const math = kit().writeEditedEquation({ ...edit, format: 'unicode', display: false, linear });
		next = (math && kit().replaceEquationPart(next, part, math)) ?? next;
	}
	return next;
}
