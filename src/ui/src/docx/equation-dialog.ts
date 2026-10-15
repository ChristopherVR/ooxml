import type { EditorView } from 'prosemirror-view';
import { checkbox, dialogButton, labelled } from './dialog-fields';
import {
	aloneOnItsLine,
	equationMarkup,
	insertEquation,
	selectedEquation,
	updateEquation,
} from './equation-commands';
import { equationDom } from './equation-schema';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

/**
 * Insert > Equation: type LaTeX (or a Word linear-format equation that uses the same commands)
 * and watch the equation take shape. With an equation selected the dialog edits it.
 */
export function createEquationDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell('Equation', 'dve-format-dialog dve-equation-dialog');
	const source = document.createElement('textarea');
	source.rows = 3;
	source.spellcheck = false;
	source.className = 'dve-equation-source';
	const display = checkbox('Display equation on its own line');
	const preview = document.createElement('div');
	preview.className = 'dve-equation-preview';
	preview.setAttribute('role', 'img');
	const message = document.createElement('p');
	message.setAttribute('role', 'status');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const content = [
		labelled('LaTeX equation', source),
		display.wrapper,
		labelled('Preview', preview),
		message,
		dialogActions(cancel, ok),
	];
	let locale: EditorLocale = 'en';
	let editing: number | undefined;
	const say = (text: string) => {
		message.textContent = text;
		localizeElement(message, locale);
	};
	const refresh = () => {
		const latex = source.value.trim();
		const omml = latex ? equationMarkup(latex, display.input.checked) : undefined;
		preview.replaceChildren();
		if (omml) preview.append(equationDom(omml, display.input.checked));
		else if (latex) preview.textContent = '';
		say(
			!latex
				? 'Type a LaTeX equation to see a preview.'
				: omml
					? ''
					: 'This LaTeX could not be converted to an equation.',
		);
		ok.disabled = !omml;
	};
	source.addEventListener('input', refresh);
	display.input.addEventListener('change', refresh);
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(getView());
	};
	ok.addEventListener('click', () => {
		const view = getView();
		if (!view) return;
		const latex = source.value.trim();
		const result =
			editing === undefined
				? insertEquation(view, latex, display.input.checked)
				: updateEquation(view, editing, latex, display.input.checked);
		if (result === 'invalid') refresh();
		else hide();
	});
	cancel.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && !ok.disabled) ok.click();
	});
	return {
		element,
		close: hide,
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return isDialogOpen(element);
		},
		open() {
			const view = getView();
			if (!view?.editable) return;
			const selected = selectedEquation(view.state);
			editing = selected?.pos;
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			source.value = selected?.latex ?? '';
			display.input.checked = selected?.display ?? aloneOnItsLine(view.state);
			refresh();
			setDialogOpen(element, true);
			source.focus();
		},
	};
}
