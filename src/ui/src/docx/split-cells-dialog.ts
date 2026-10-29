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
import type { EditorLocale } from './localization';
import { executeTableCommand } from './table-commands';
import { splitTarget } from './table-restructure-commands';

/** The Split Cells dialog; OK splits the cell around the selection into the chosen grid. */
export function createSplitCellsDialog(viewOf: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell('Split cells', 'dve-format-dialog');
	const columns = numberInput(1, 63, 1);
	const rows = numberInput(1, 63, 1);
	const note = document.createElement('p');
	note.textContent = 'Split the cell into this many columns and rows.';
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const content = [
		fieldset(
			'Split cells',
			row(labelled('Number of columns', columns), labelled('Number of rows', rows)),
		),
		note,
		dialogActions(cancel, ok),
	];
	let locale: EditorLocale = 'en';
	const valid = () =>
		[columns, rows].every(
			(input) => Number.isInteger(Number(input.value)) && Number(input.value) >= 1,
		);
	const validate = () => (ok.disabled = !valid());
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		viewOf()?.focus();
	};
	const submit = () => {
		const view = viewOf();
		if (!view || !valid()) return;
		const split = { rows: Number(rows.value), columns: Number(columns.value) };
		hide();
		executeTableCommand(view, 'splitCell', split);
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
			columns.value = '2';
			rows.value = '1';
			validate();
			setDialogOpen(element, true);
			columns.focus();
		},
	};
}
