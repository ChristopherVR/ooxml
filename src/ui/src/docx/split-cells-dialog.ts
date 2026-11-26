// Table > Split Cells: how many columns and rows the selected cell becomes.
import type { EditorView } from 'prosemirror-view';
import { dialogButton, fieldset, labelled, numberInput, row } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import type { FormatDialog } from './font-dialog';
import { translate, type EditorLocale } from './localization';
import { canExecuteTableCommand, executeTableCommand } from './table-commands';
import { splitTarget } from './table-restructure-commands';

/** The Split Cells dialog; OK splits the cell around the selection into the chosen grid. */
export function createSplitCellsDialog(viewOf: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell('Split cells', 'dve-format-dialog');
	const columns = numberInput(1, 63, 1);
	const rows = numberInput(1, 63, 1);
	const note = document.createElement('p');
	note.textContent = 'Split the cell into this many columns and rows.';
	const error = document.createElement('p');
	error.setAttribute('role', 'alert');
	error.hidden = true;
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const content = [
		fieldset(
			'Split cells',
			row(labelled('Number of columns', columns), labelled('Number of rows', rows)),
		),
		note,
		error,
		dialogActions(cancel, ok),
	];
	let locale: EditorLocale = 'en';
	const valid = () =>
		[columns, rows].every(
			(input) => input.checkValidity() && Number.isInteger(input.valueAsNumber),
		);
	const dimensions = () => ({ rows: rows.valueAsNumber, columns: columns.valueAsNumber });
	const validate = () => {
		const view = viewOf();
		const sizeValid = valid();
		const supported = view && sizeValid && canExecuteTableCommand(view, 'splitCell', dimensions());
		error.hidden = !sizeValid || Boolean(supported);
		ok.disabled = !supported;
		if (error.hidden) ok.removeAttribute('aria-description');
		else ok.setAttribute('aria-description', error.textContent ?? '');
	};
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		viewOf()?.focus();
	};
	const submit = () => {
		const view = viewOf();
		if (!view || !valid()) return;
		if (executeTableCommand(view, 'splitCell', dimensions())) hide();
		else validate();
	};
	columns.addEventListener('input', validate);
	rows.addEventListener('input', validate);
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', submit);
	onDialogDismiss(element, hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
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
			const view = viewOf();
			const target = view && splitTarget(view);
			if (!view?.editable || !target) return;
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			error.textContent = translate(locale, 'tableTab.splitDimensions');
			columns.value = String(target.columns);
			rows.value = String(target.rows);
			validate();
			setDialogOpen(element, true);
			columns.focus();
		},
	};
}
