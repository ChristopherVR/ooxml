import type { VisioMasterFormat, VisioShape } from 'ooxml-core/visio';
import { ViewerDialog } from './viewer-dialog';

type Translate = (text: string) => string;
let nextField = 0;

const hex = (color: string | undefined) =>
	color && /^#[0-9a-f]{6}$/i.test(color) ? color.toLowerCase() : '#000000';
const round = (value: number, places = 3) => String(Number(value.toFixed(places)));

/**
 * Edit Master: the master shape's default text, size, fill, line and text formatting in one
 * dialog. Resolves with only what the user changed (nothing changed resolves with `{}`), or
 * `undefined` when the dialog is cancelled. `note` is shown under the fields.
 */
export function askMasterFormat(
	pane: HTMLElement,
	shape: VisioShape,
	t: Translate,
	note?: string,
): Promise<VisioMasterFormat | undefined> {
	const doc = pane.ownerDocument;
	return new Promise((resolve) => {
		const ok = t('OK');
		const form = doc.createElement('form');
		const field = (label: string, input: HTMLElement) => {
			const wrapper = doc.createElement('label');
			wrapper.textContent = t(label);
			input.id = `stencil-master-field-${++nextField}`;
			wrapper.append(input);
			form.append(wrapper);
			return input;
		};
		const text = field('Default shape text', doc.createElement('textarea')) as HTMLTextAreaElement;
		text.maxLength = 4096;
		text.value = shape.text.plainText;
		const number = (label: string, value: number, step: string, min: string) => {
			const input = field(label, doc.createElement('input')) as HTMLInputElement;
			input.type = 'number';
			input.step = step;
			input.min = min;
			input.value = round(value);
			return input;
		};
		const color = (label: string, value: string | undefined) => {
			const input = field(label, doc.createElement('input')) as HTMLInputElement;
			input.type = 'color';
			input.value = hex(value);
			return input;
		};
		const width = number('Width (in)', shape.width, '0.01', '0.01');
		const height = number('Height (in)', shape.height, '0.01', '0.01');
		const fill = color('Fill colour', shape.style.fill);
		const line = color('Line colour', shape.style.lineColor);
		const weight = number('Line weight (pt)', shape.style.lineWidth * 72, '0.25', '0');
		const size = number('Font size (pt)', shape.text.fontSize * 72, '0.5', '1');
		const fontColor = color('Text colour', shape.text.color);
		const flags = (['Bold', 'Italic', 'Underline'] as const).map((label) => {
			const wrapper = doc.createElement('label');
			const input = doc.createElement('input');
			input.type = 'checkbox';
			input.checked = Boolean(shape.text[label.toLowerCase() as 'bold' | 'italic' | 'underline']);
			wrapper.append(input, ` ${t(label)}`);
			form.append(wrapper);
			return input;
		});
		const align = field('Text alignment', doc.createElement('select')) as HTMLSelectElement;
		for (const [value, label] of [
			['', 'Unchanged'],
			['left', 'Left'],
			['center', 'Center'],
			['right', 'Right'],
			['justify', 'Justify'],
		] as const) {
			const option = doc.createElement('option');
			option.value = value;
			option.textContent = t(label);
			align.append(option);
		}
		if (note) {
			const paragraph = doc.createElement('p');
			paragraph.textContent = t(note);
			form.append(paragraph);
		}
		const initial = {
			fill: fill.value,
			line: line.value,
			fontColor: fontColor.value,
			flags: flags.map((input) => input.checked),
		};
		form.addEventListener('submit', (event) => {
			event.preventDefault();
			const format: VisioMasterFormat = {};
			if (text.value !== shape.text.plainText) format.text = text.value;
			const changedNumber = (input: HTMLInputElement, before: number) => {
				const value = Number(input.value);
				return input.value !== round(before) && Number.isFinite(value) ? value : undefined;
			};
			const w = changedNumber(width, shape.width);
			const h = changedNumber(height, shape.height);
			const lw = changedNumber(weight, shape.style.lineWidth * 72);
			const fs = changedNumber(size, shape.text.fontSize * 72);
			if (w !== undefined) format.width = w;
			if (h !== undefined) format.height = h;
			if (lw !== undefined) format.lineWeight = lw;
			if (fs !== undefined) format.fontSize = fs;
			if (fill.value !== initial.fill) format.fillColor = fill.value;
			if (line.value !== initial.line) format.lineColor = line.value;
			if (fontColor.value !== initial.fontColor) format.fontColor = fontColor.value;
			(['bold', 'italic', 'underline'] as const).forEach((key, index) => {
				if (flags[index]!.checked !== initial.flags[index]) format[key] = flags[index]!.checked;
			});
			if (align.value) format.horizontalAlign = align.value as 'left';
			resolve(format);
			dialog.close();
		});
		const dialog = new ViewerDialog(
			pane,
			'stencil-dialog',
			t('Edit Master'),
			[ok, t('Cancel')],
			(button) => {
				if (button === ok) form.requestSubmit();
				else dialog.close();
			},
		);
		dialog.body.append(form);
		dialog.dialog.setAttribute('close-label', t('Close'));
		dialog.dialog.addEventListener(
			'office-dialog-close',
			() => {
				resolve(undefined);
				queueMicrotask(() => dialog.dialog.remove());
			},
			{ once: true },
		);
		dialog.show();
	});
}
