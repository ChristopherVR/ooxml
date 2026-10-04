import type { WatermarkSpec } from 'docx-core';
import {
	checkbox,
	dialogButton,
	labelled,
	listInput,
	row,
	selectOf,
	textInput,
} from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

export interface WatermarkHost {
	current(): WatermarkSpec | undefined;
	canEdit(): boolean;
	/** Sets the watermark in every header; undefined removes it. */
	apply(spec: WatermarkSpec | undefined): void;
	restoreFocus(): void;
}

const PRESETS = ['CONFIDENTIAL', 'DRAFT', 'DO NOT COPY', 'SAMPLE', 'URGENT'];
const DEFAULT_COLOR = '#c0c0c0';

/** Word's Printed Watermark for text: the words, font, colour, layout and transparency. */
export function createWatermarkDialog(host: WatermarkHost): FormatDialog {
	const element = createDialogShell('Printed watermark', 'dve-format-dialog dve-watermark-dialog');
	const setting = selectOf([
		['none', 'No watermark'],
		['text', 'Text watermark'],
	]);
	const { input: text, list } = listInput(PRESETS);
	const font = textInput();
	const color = document.createElement('input');
	color.type = 'color';
	const layout = selectOf([
		['diagonal', 'Diagonal'],
		['horizontal', 'Horizontal'],
	]);
	const semi = checkbox('Semitransparent');
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = dialogActions(cancel, ok);
	const content = [
		labelled('Setting', setting),
		labelled('Text', text),
		list,
		row(labelled('Font', font), labelled('Color', color)),
		row(labelled('Layout', layout), semi.wrapper),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	const validate = () => {
		const on = setting.value === 'text';
		for (const control of [text, font, color, layout, semi.input]) control.disabled = !on;
		const valid = !on || text.value.trim() !== '';
		message.textContent = valid ? '' : 'Enter the watermark text.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	for (const control of [setting, text, font, color, layout, semi.input])
		for (const type of ['input', 'change']) control.addEventListener(type, validate);
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		if (!host.canEdit() || !validate()) return;
		host.apply(
			setting.value === 'text'
				? {
						text: text.value.trim(),
						color: color.value.toLowerCase(),
						semitransparent: semi.input.checked,
						layout: layout.value as WatermarkSpec['layout'],
						...(font.value.trim() ? { fontFamily: font.value.trim() } : {}),
					}
				: undefined,
		);
		hide();
	};
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
			if (!host.canEdit()) return;
			const current = host.current();
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			setting.value = current ? 'text' : 'none';
			text.value = current?.text ?? '';
			font.value = current?.fontFamily ?? 'Calibri';
			color.value = current?.color ?? DEFAULT_COLOR;
			layout.value = current?.layout ?? 'diagonal';
			semi.input.checked = current?.semitransparent ?? true;
			validate();
			setDialogOpen(element, true);
			setting.focus();
		},
	};
}
