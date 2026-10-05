import type { DocumentModel } from 'ooxml-core/docx';
import type { EditorView } from 'prosemirror-view';
import { CAPTION_LABELS, insertCaption } from './caption-commands';
import { dialogButton, labelled, selectOf, textInput } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { translate, type EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

/** The document's Caption style id, matched by id or name. */
function captionStyle(model: DocumentModel): string | undefined {
	return Object.values(model.paragraphStyles?.styles ?? {}).find(
		(style) => /^caption$/i.test(style.id) || /^caption$/i.test(style.name ?? ''),
	)?.id;
}

/** Word's Caption dialog: a label, its numbered `SEQ` field and optional text, above or below. */
export function createCaptionDialog(
	getView: () => EditorView | undefined,
	getModel: () => DocumentModel,
): FormatDialog {
	const element = createDialogShell('Insert caption', 'dve-format-dialog dve-caption-dialog');
	const label = selectOf(CAPTION_LABELS.map((name) => [name, name] as const));
	const text = textInput();
	const position = selectOf([
		['below', 'Below selected item'],
		['above', 'Above selected item'],
	]);
	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('OK', true);
	const actions = dialogActions(cancel, confirm);
	element.append(
		labelled('Label', label),
		labelled('Caption', text),
		labelled('Position', position),
		actions,
	);
	let locale: EditorLocale = 'en';
	const content = [...element.childNodes];
	element.replaceChildren();
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(getView());
	};
	// Tables are captioned above and everything else below, as Word defaults.
	label.addEventListener('change', () => {
		position.value = label.value === 'Table' ? 'above' : 'below';
	});
	const submit = () => {
		const view = getView();
		const style = captionStyle(getModel());
		if (view)
			insertCaption(view, {
				label: label.value,
				labelText: translate(locale, label.value as never),
				text: text.value,
				position: position.value === 'above' ? 'above' : 'below',
				...(style ? { style } : {}),
			});
		hide();
	};
	cancel.addEventListener('click', hide);
	confirm.addEventListener('click', submit);
	onDialogDismiss(element, hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});
	return {
		element,
		open() {
			if (!getView()?.editable) return;
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			label.value = 'Figure';
			position.value = 'below';
			text.value = '';
			setDialogOpen(element, true);
			text.focus();
		},
		close: hide,
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return isDialogOpen(element);
		},
	};
}
