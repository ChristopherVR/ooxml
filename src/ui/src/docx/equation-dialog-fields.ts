// The Insert or Edit Equation dialog's contents (LaTeX source, display choice and live preview),
// loaded when the dialog first opens.
import type { EditorView } from 'prosemirror-view';
import type { Selection } from 'prosemirror-state';
import { createEquationFields } from '../math/equation-fields';
import {
	aloneOnItsLine,
	insertEquation,
	selectedEquation,
	updateEquation,
} from './equation-commands';
import { localizeDialog, setDialogOpen } from './dialog-shell';
import { translate, type EditorLocale } from './localization';

export interface EquationDialogParts {
	element: HTMLElement;
	ok: HTMLButtonElement;
	actions: HTMLElement;
	locale: EditorLocale;
	view: EditorView;
	selection: Selection;
	getView(): EditorView | undefined;
	hide(): void;
}

/** Fills the dialog for `view`'s selection and opens it; returns what OK runs. */
export function showEquationFields(parts: EquationDialogParts): () => void {
	const { element, ok, actions, locale, view, selection, getView, hide } = parts;
	const selected = selectedEquation(view.state);
	const fields = createEquationFields(element.ownerDocument, {
		labels: {
			equation: translate(locale, 'Equation'),
			source: translate(locale, 'LaTeX equation'),
			display: translate(locale, 'Display equation on its own line'),
			preview: translate(locale, 'Preview'),
			empty: translate(locale, 'Type a LaTeX equation to see a preview.'),
			invalid: translate(locale, 'This LaTeX could not be converted to an equation.'),
		},
		...(selected ? { latex: selected.latex } : {}),
		display: selected?.display ?? aloneOnItsLine(view.state),
		allowInline: true,
		onChange: (value) => {
			ok.disabled = !value;
		},
		accept: () => ok.click(),
	});
	fields.source.classList.add('dve-equation-source');
	fields.preview.classList.add('dve-equation-preview');
	element.replaceChildren(fields.element, actions);
	localizeDialog(element, locale);
	setDialogOpen(element, true);
	fields.source.focus();
	return () => {
		const value = fields.draft;
		if (!value || getView() !== view || !selection.eq(view.state.selection)) return;
		if (!selected || value.latex !== selected.latex || value.display !== selected.display) {
			if (selected) updateEquation(view, selected.pos, value.latex, value.display);
			else insertEquation(view, value.latex, value.display);
		}
		hide();
	};
}
