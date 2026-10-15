import type { EditorView } from 'prosemirror-view';
import { dialogButton, labelled } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import { selectedShape, setShapeText } from './shape-commands';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

/** Shape Format > Add Text: the text inside the selected shape, one paragraph per line. */
export function createShapeTextDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell('Add Text', 'dve-format-dialog dve-shape-text-dialog');
	const text = document.createElement('textarea');
	text.rows = 5;
	const message = document.createElement('p');
	message.setAttribute('role', 'status');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const content = [labelled('Text', text), message, dialogActions(cancel, ok)];
	let locale: EditorLocale = 'en';
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(getView());
	};
	cancel.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	ok.addEventListener('click', () => {
		const view = getView();
		if (view) setShapeText(view, text.value.split(/\r\n|\r|\n/));
		hide();
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
			const shape = view && selectedShape(view.state);
			if (!view?.editable) return;
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			const editable = Boolean(shape?.spec.textEditable);
			text.value = shape?.spec.text?.join('\n') ?? '';
			text.disabled = !editable;
			ok.disabled = !editable;
			message.textContent = !shape
				? 'Select a shape first.'
				: editable
					? ''
					: 'This shape has formatted text that the editor cannot change.';
			localizeElement(message, locale);
			setDialogOpen(element, true);
			text.focus();
		},
	};
}
