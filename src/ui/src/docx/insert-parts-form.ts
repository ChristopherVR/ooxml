import type { Block, GlossaryEntry, SignatureLineSpec, WebVideoSpec } from 'ooxml-core/docx';
import type { DocumentPropertySet } from 'ooxml-core/opc';
import type { FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';

/** What the Quick Parts, Signature Line and Online Video dialogs need from the editor. */
export interface InsertPartsHost {
	editable(): boolean;
	focus(): void;
	glossary(): readonly GlossaryEntry[];
	/** Adds an entry, replacing one with the same name, gallery and category. */
	saveGlossaryEntry(entry: GlossaryEntry): void;
	removeGlossaryEntry(key: string): void;
	insertEntry(entry: GlossaryEntry): void;
	selectionBlocks(): Block[] | undefined;
	/** Building blocks are written to the glossary part, which a shared document does not take. */
	glossaryLocked(): boolean;
	/** Tells the user why a command did nothing. */
	report(message: string): void;
	insertSignatureLine(spec: SignatureLineSpec): void;
	insertWebVideo(spec: WebVideoSpec): void;
	onlineVideoHosts(): readonly string[];
	properties(): DocumentPropertySet;
	insertField(instr: string, result: string): void;
}

export interface Form {
	title: string;
	className: string;
	content: HTMLElement[];
	ok: HTMLButtonElement;
	cancel: HTMLButtonElement;
	/** Fills the fields; returns false when the dialog cannot open now. */
	reset(locale: EditorLocale): boolean;
	/** Whether the OK button is enabled. */
	valid(): boolean;
	/** Applies the dialog; true closes it. */
	apply(): boolean;
	/** The control that takes focus when the dialog opens. */
	first: HTMLElement;
	/** Fields whose change re-validates. */
	inputs: HTMLElement[];
}

/** One modal dialog from a form spec: shell, localisation, dismissal and the OK and Cancel flow. */
export function formDialog(host: InsertPartsHost, form: Form): FormatDialog {
	const element = createDialogShell(form.title, `dve-format-dialog ${form.className}`);
	let locale: EditorLocale = 'en';
	const actions = dialogActions(form.cancel, form.ok);
	let origin: HTMLElement | undefined;
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		// Focus goes back to the control that opened the dialog, else to the document.
		if (origin?.isConnected) origin.focus();
		else host.focus();
	};
	const validate = () => {
		form.ok.disabled = !form.valid();
	};
	for (const input of form.inputs) {
		input.addEventListener('input', validate);
		input.addEventListener('change', validate);
	}
	form.cancel.addEventListener('click', hide);
	form.ok.addEventListener('click', () => {
		if (host.editable() && form.valid() && form.apply()) hide();
	});
	onDialogDismiss(element, hide);
	element.addEventListener('keydown', (event) => {
		// Escape closes the dialog and nothing behind it (a folded ribbon group, say).
		if (event.key === 'Escape') event.stopPropagation();
		else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			if (!form.ok.disabled) form.ok.click();
		}
	});
	return {
		element,
		close: hide,
		get isOpen() {
			return isDialogOpen(element);
		},
		setLocale(next) {
			locale = next;
		},
		open(opener) {
			origin = opener;
			if (!host.editable() || !form.reset(locale)) return;
			element.replaceChildren(...form.content, actions);
			localizeDialog(element, locale);
			validate();
			setDialogOpen(element, true);
			form.first.focus();
		},
	};
}

export function note(text: string): HTMLParagraphElement {
	const paragraph = document.createElement('p');
	paragraph.textContent = text;
	return paragraph;
}
