import type { EditorCore } from './editor-core';
import {
	dialogActions,
	createDialogShell,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import { dialogButton, labelled, numberInput, selectOf, textInput } from './dialog-fields';
import { translate, type LocalizationKey } from './localization';
import type { TableSort, TableTextSeparator } from 'ooxml-core/docx/ui';

function open(
	core: EditorCore,
	title: LocalizationKey,
	fields: HTMLElement[],
	submit: () => boolean,
	invalid: LocalizationKey = 'tableTab.valueError',
): void {
	const dialog = createDialogShell(title, 'dve-format-dialog');
	const error = document.createElement('p');
	error.setAttribute('role', 'alert');
	error.hidden = true;
	error.textContent = translate(core.locale, invalid);
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	dialog.append(...fields, error, dialogActions(cancel, ok));
	const close = () => {
		setDialogOpen(dialog, false);
		dialog.remove();
	};
	cancel.addEventListener('click', close);
	onDialogDismiss(dialog, close);
	ok.addEventListener('click', () => {
		if (submit()) close();
		else error.hidden = false;
	});
	localizeDialog(dialog, core.locale);
	const root = core.shell.toolbar?.getRootNode();
	if (root instanceof ShadowRoot) root.append(dialog);
	else core.element.append(dialog);
	setDialogOpen(dialog, true);
}

export function openTableFormula(
	core: EditorCore,
	above: boolean,
	submit: (value: string) => boolean,
): void {
	const input = textInput();
	input.maxLength = 128;
	input.value = above ? '=SUM(ABOVE)' : '=SUM(LEFT)';
	open(
		core,
		'Formula',
		[labelled('Formula', input)],
		() => submit(input.value),
		'tableTab.formulaError',
	);
}

export function openTableConvert(
	core: EditorCore,
	submit: (value: TableTextSeparator) => boolean,
): void {
	const select = selectOf([
		['tabs', 'Tabs'],
		['commas', 'Commas'],
		['paragraphs', 'Paragraph marks'],
	]);
	open(core, 'Convert to Text', [labelled('Separate text with', select)], () =>
		submit(select.value === 'commas' || select.value === 'paragraphs' ? select.value : 'tabs'),
	);
}

export function openTableSort(
	core: EditorCore,
	width: number,
	column: number,
	headerRows: number,
	submit: (value: TableSort) => boolean,
): void {
	const key = numberInput(1, width, 1);
	key.valueAsNumber = column + 1;
	const order = selectOf([
		['ascending', 'Ascending'],
		['descending', 'Descending'],
	]);
	const type = selectOf([
		['text', 'Text'],
		['number', 'Number'],
	]);
	const header = selectOf([
		['no', 'No header row'],
		['yes', 'Header Row'],
	]);
	header.value = headerRows ? 'yes' : 'no';
	open(
		core,
		'Sort Table',
		[
			labelled('Sort Column', key),
			labelled('Sort Order', order),
			labelled('Sort Type', type),
			labelled('Sort Headers', header),
		],
		() =>
			key.checkValidity() &&
			Number.isSafeInteger(key.valueAsNumber) &&
			submit({
				column: key.valueAsNumber - 1,
				order: order.value === 'descending' ? 'descending' : 'ascending',
				type: type.value === 'number' ? 'number' : 'text',
				headerRows: header.value === 'yes' ? Math.max(1, headerRows) : 0,
				locale: core.locale,
			}),
	);
}
