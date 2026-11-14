// Home > Paste > Paste Special (Ctrl+Alt+V): choose which of the clipboard's formats to paste.
import type { EditorView } from 'prosemirror-view';
import { readClipboardContents, type ClipboardContents } from './context-menu-actions';
import { dialogButton, labelled } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import type { FormatDialog } from './font-dialog';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import { pastePayload, type PasteOption } from './paste-options';

/** The formats Paste Special offers for what the browser let the editor read. */
export function pasteSpecialFormats(
	contents: ClipboardContents,
): Array<{ option: PasteOption; label: LocalizationKey }> {
	return [
		...(contents.html ? [{ option: 'source' as const, label: 'HTML Format' as const }] : []),
		...(contents.text !== null
			? [{ option: 'text' as const, label: 'Unformatted Text' as const }]
			: []),
	];
}

export function createPasteSpecialDialog(viewOf: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell('Paste Special', 'dve-format-dialog dve-paste-special');
	const formats = document.createElement('select');
	formats.size = 3;
	const note = document.createElement('p');
	note.className = 'dve-paste-special-note';
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	let locale: EditorLocale = 'en';
	let contents: ClipboardContents | null = null;
	let opening = 0;
	const hide = () => {
		opening++;
		setDialogOpen(element, false);
		element.replaceChildren();
		viewOf()?.focus();
	};
	const submit = () => {
		const view = viewOf();
		const option = formats.value as PasteOption;
		if (!view || !contents || !formats.value) return;
		const payload = { html: contents.html, text: contents.text ?? '' };
		hide();
		pastePayload(view, payload, option);
	};
	const show = (message: LocalizationKey | null) => {
		const content = message
			? [note, dialogActions(ok)]
			: [labelled('Paste as:', formats), dialogActions(cancel, ok)];
		element.replaceChildren(...content);
		localizeDialog(element, locale);
		note.textContent = message ? translate(locale, message) : '';
		setDialogOpen(element, true);
		(message ? ok : formats).focus();
	};
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', () => (contents && formats.value ? submit() : hide()));
	formats.addEventListener('dblclick', submit);
	formats.addEventListener('keydown', (event) => {
		if (event.key !== 'Enter') return;
		event.preventDefault();
		submit();
	});
	onDialogDismiss(element, hide);
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
			if (!view?.editable) return;
			const attempt = ++opening;
			contents = null;
			void readClipboardContents()
				.then((read) => (read ? { read } : { denied: true }))
				.catch(() => ({ denied: true }))
				.then((result) => {
					if (attempt !== opening) return;
					if ('denied' in result) {
						show(
							'The browser did not let the editor read the clipboard. Press Ctrl+V to paste instead.',
						);
						return;
					}
					const offered = pasteSpecialFormats(result.read);
					if (!offered.length) {
						show('The clipboard holds nothing the editor can paste.');
						return;
					}
					contents = result.read;
					formats.replaceChildren(
						...offered.map(({ option, label }) => new Option(translate(locale, label), option)),
					);
					formats.selectedIndex = 0;
					show(null);
				});
		},
	};
}
