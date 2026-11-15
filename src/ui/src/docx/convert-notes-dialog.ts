import type { NoteConversion } from 'ooxml-core/docx';
import { dialogButton, fieldset } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

export interface ConvertNotesHost {
	/** How many footnotes and endnotes the document has. */
	counts(): { footnotes: number; endnotes: number };
	canEdit(): boolean;
	apply(conversion: NoteConversion): void;
	restoreFocus(): void;
}

const CHOICES: [NoteConversion, string][] = [
	['footnotesToEndnotes', 'Convert all footnotes to endnotes'],
	['endnotesToFootnotes', 'Convert all endnotes to footnotes'],
	['swap', 'Swap footnotes and endnotes'],
];

/** Word's Convert Notes: every footnote to an endnote, every endnote to a footnote, or a swap. */
export function createConvertNotesDialog(host: ConvertNotesHost): FormatDialog {
	const element = createDialogShell('Convert Notes', 'dve-format-dialog dve-convert-notes-dialog');
	const radios = new Map<NoteConversion, HTMLInputElement>();
	const options = CHOICES.map(([value, label]) => {
		const wrapper = document.createElement('label');
		wrapper.className = 'dve-dialog-check';
		const input = document.createElement('input');
		input.type = 'radio';
		input.name = 'dve-convert-notes';
		input.value = value;
		input.setAttribute('aria-label', label);
		const text = document.createElement('span');
		text.textContent = label;
		wrapper.append(input, text);
		radios.set(value, input);
		return wrapper;
	});
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const content = [fieldset('Convert Notes', ...options), dialogActions(cancel, ok)];
	let locale: EditorLocale = 'en';
	const chosen = () => [...radios].find(([, input]) => input.checked && !input.disabled)?.[0];
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		const value = chosen();
		if (!host.canEdit() || !value) return;
		host.apply(value);
		hide();
	};
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', submit);
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
			if (!host.canEdit()) return;
			const { footnotes, endnotes } = host.counts();
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			radios.get('footnotesToEndnotes')!.disabled = footnotes === 0;
			radios.get('endnotesToFootnotes')!.disabled = endnotes === 0;
			radios.get('swap')!.disabled = footnotes === 0 && endnotes === 0;
			for (const input of radios.values()) input.checked = false;
			const first = [...radios.values()].find((input) => !input.disabled);
			if (first) first.checked = true;
			ok.disabled = !first;
			setDialogOpen(element, true);
			(first ?? cancel).focus();
		},
	};
}
