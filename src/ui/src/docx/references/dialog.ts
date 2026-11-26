import { checkbox, dialogButton, labelled, selectOf, textInput } from '../dialog-fields';
import { createDialogShell, dialogActions, onDialogDismiss, setDialogOpen } from '../dialog-shell';
import type { EditorCore } from '../editor-core';
import { referencesText, type ReferencesText } from './strings';
import css from './dialog.css?raw';

export function referencesDialog(core: EditorCore, title: ReferencesText) {
	const { loadGeneration, view, locale } = core;
	const initialDoc = view?.state.doc;
	const t = (key: ReferencesText) => referencesText(locale, key);
	const frame = core.element.shadowRoot?.querySelector('.dve-frame');
	frame?.querySelector<HTMLElement>('.dve-references-dialog [data-dismiss]')?.click();
	const element = createDialogShell(t(title), 'dve-references-dialog');
	element.setAttribute('close-label', t('Close'));
	const style = document.createElement('style');
	style.textContent = css;
	const content = document.createElement('div');
	content.className = 'dve-references-content';
	const status = document.createElement('p');
	status.setAttribute('role', 'status');
	status.className = 'dve-references-status';
	const close = () => {
		setDialogOpen(element, false);
		element.remove();
	};
	const cancel = dialogButton(t('Cancel'));
	cancel.dataset.dismiss = '';
	cancel.addEventListener('click', close);
	onDialogDismiss(element, close);
	element.append(style, content, status);
	return {
		t,
		content,
		text(label: ReferencesText, value = '') {
			const input = textInput();
			input.maxLength = 512;
			input.value = value;
			content.append(labelled(t(label), input));
			return input;
		},
		select<T extends string>(
			label: ReferencesText,
			options: readonly (readonly [T, string])[],
			value?: T,
		) {
			const input = selectOf(options);
			if (value !== undefined) input.value = value;
			content.append(labelled(t(label), input));
			return {
				input,
				value: () => options.find(([key]) => key === input.value)?.[0],
			};
		},
		check(label: ReferencesText, checked = false) {
			const { input, wrapper } = checkbox(t(label));
			input.checked = checked;
			content.append(wrapper);
			return input;
		},
		action(label: ReferencesText, apply: () => void, primary = false) {
			const button = dialogButton(t(label), primary);
			button.addEventListener('click', () => {
				if (
					loadGeneration !== core.loadGeneration ||
					view !== core.view ||
					!view?.editable ||
					initialDoc !== view.state.doc
				) {
					status.textContent = t('references.changed');
					return;
				}
				try {
					apply();
					close();
				} catch (error) {
					status.textContent = error instanceof Error ? error.message : String(error);
				}
			});
			return button;
		},
		show(...buttons: HTMLElement[]) {
			element.append(dialogActions(cancel, ...buttons));
			frame?.append(element);
			setDialogOpen(element, true);
		},
	};
}

export type ReferencesDialog = ReturnType<typeof referencesDialog>;
