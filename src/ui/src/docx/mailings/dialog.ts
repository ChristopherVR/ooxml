import { dialogButton } from '../dialog-fields';
import { createDialogShell, dialogActions, onDialogDismiss, setDialogOpen } from '../dialog-shell';
import { translate, type EditorLocale, type LocalizationKey } from '../localization';
import { mailingsText, type MailingsLabel } from './strings';
import css from './dialog.css?raw';

export interface MailingsDialog {
	element: HTMLElement;
	content: HTMLElement;
	t: (label: MailingsLabel) => string;
	close(): void;
	actions(...buttons: HTMLElement[]): void;
	show(): void;
}

export function createMailingsDialog(
	host: ShadowRoot,
	title: LocalizationKey,
	locale: EditorLocale,
): MailingsDialog {
	const element = createDialogShell(translate(locale, title), 'dve-mailings-dialog');
	element.setAttribute('close-label', translate(locale, 'Close'));
	const style = document.createElement('style');
	style.textContent = css;
	const content = document.createElement('div');
	content.className = 'dve-mailings-content';
	element.append(style, content);
	const close = () => {
		setDialogOpen(element, false);
		element.remove();
	};
	onDialogDismiss(element, close);
	return {
		element,
		content,
		t: (label) => mailingsText(locale, label),
		close,
		actions(...buttons) {
			const dismiss = dialogButton(translate(locale, buttons.length ? 'Cancel' : 'Close'));
			dismiss.addEventListener('click', close);
			element.append(dialogActions(dismiss, ...buttons));
		},
		show() {
			host.querySelector('.dve-frame')?.append(element);
			setDialogOpen(element, true);
		},
	};
}

export function dialogMessage(dialog: MailingsDialog, text: MailingsLabel): HTMLParagraphElement {
	const message = document.createElement('p');
	message.textContent = dialog.t(text);
	dialog.content.append(message);
	return message;
}

export function dialogChoice(
	dialog: MailingsDialog,
	label: MailingsLabel,
	run: () => void,
	reason?: MailingsLabel,
): HTMLButtonElement {
	const button = dialogButton(dialog.t(label));
	if (reason) {
		button.disabled = true;
		button.title = dialog.t(reason);
		dialogMessage(dialog, reason);
	}
	button.addEventListener('click', run);
	return button;
}

export function dialogSubmit(dialog: MailingsDialog, locale: EditorLocale, run: () => void) {
	const ok = dialogButton(translate(locale, 'OK'), true);
	ok.addEventListener('click', run);
	dialog.actions(ok);
	return ok;
}
