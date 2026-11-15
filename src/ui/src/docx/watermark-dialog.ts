import {
	PICTURE_WATERMARK_SCALES,
	pictureWatermarkSizePt,
	type PictureWatermarkSpec,
} from 'ooxml-core/docx';
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
import { isPictureWatermark, type AnyWatermark } from './watermark-commands';

/** A picture chosen for a watermark, staged for saving, with its natural size in points. */
export interface WatermarkPicture {
	partName: string;
	contentType: string;
	widthPt: number;
	heightPt: number;
	title?: string;
}

export interface WatermarkHost {
	current(): AnyWatermark | undefined;
	canEdit(): boolean;
	/** Sets the watermark in every header; undefined removes it. */
	apply(spec: AnyWatermark | undefined): void;
	/** Stages a chosen picture file; rejects for a format Word pictures cannot use. */
	stagePicture?(file: File): Promise<WatermarkPicture>;
	/** The natural size of a picture already in the document, for rescaling it. */
	pictureSize?(
		partName: string,
		contentType: string,
	): Promise<{ widthPt: number; heightPt: number } | undefined>;
	/** The text area Auto scale fits the picture to, in points. */
	textArea?(): { widthPt: number; heightPt: number };
	restoreFocus(): void;
	reportError?(cause: unknown): void;
}

const PRESETS = ['CONFIDENTIAL', 'DRAFT', 'DO NOT COPY', 'SAMPLE', 'URGENT'];
const DEFAULT_COLOR = '#c0c0c0';
const DEFAULT_AREA = { widthPt: 468, heightPt: 648 };

/** Word's Printed Watermark: none, a picture (scale and washout), or text (font, colour, layout). */
export function createWatermarkDialog(host: WatermarkHost): FormatDialog {
	const element = createDialogShell('Printed watermark', 'dve-format-dialog dve-watermark-dialog');
	const setting = selectOf([
		['none', 'No watermark'],
		['picture', 'Picture watermark'],
		['text', 'Text watermark'],
	]);
	const choose = dialogButton('Select Picture…');
	const fileInput = document.createElement('input');
	fileInput.type = 'file';
	fileInput.accept = 'image/png,image/jpeg,image/gif,image/bmp';
	fileInput.hidden = true;
	const pictureName = document.createElement('span');
	pictureName.className = 'dve-watermark-picture-name';
	const scale = selectOf([
		['auto', 'Auto'],
		...PICTURE_WATERMARK_SCALES.map((value) => [String(value), `${value}%`] as const),
	]);
	const washout = checkbox('Washout');
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
		row(choose, pictureName),
		fileInput,
		row(labelled('Scale', scale), washout.wrapper),
		labelled('Text', text),
		list,
		row(labelled('Font', font), labelled('Color', color)),
		row(labelled('Layout', layout), semi.wrapper),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	/** The picture to use; `natural` marks a newly chosen one whose size is its own. */
	let picture: (WatermarkPicture & { natural?: boolean }) | undefined;
	let initial: PictureWatermarkSpec | undefined;
	const validate = () => {
		const mode = setting.value;
		for (const control of [text, font, color, layout, semi.input])
			control.disabled = mode !== 'text';
		for (const control of [choose, scale, washout.input]) control.disabled = mode !== 'picture';
		const valid =
			mode === 'none' || (mode === 'text' ? text.value.trim() !== '' : Boolean(picture));
		message.textContent = valid
			? ''
			: mode === 'text'
				? 'Enter the watermark text.'
				: 'Select a picture for the watermark.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	for (const control of [setting, text, font, color, layout, semi.input, scale, washout.input])
		for (const type of ['input', 'change']) control.addEventListener(type, validate);
	choose.addEventListener('click', () => fileInput.click());
	fileInput.addEventListener('change', async () => {
		const file = fileInput.files?.[0];
		fileInput.value = '';
		if (!file || !host.stagePicture) return;
		try {
			picture = { ...(await host.stagePicture(file)), natural: true };
			pictureName.textContent = file.name;
		} catch (cause) {
			host.reportError?.(cause);
		}
		validate();
	});
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		host.restoreFocus();
	};
	const pictureSpec = async (): Promise<PictureWatermarkSpec | undefined> => {
		if (!picture) return undefined;
		const chosen = scale.value === 'auto' ? 'auto' : Number(scale.value);
		const unchanged =
			initial &&
			picture.partName === initial.partName &&
			(initial.scale ?? 'auto') === chosen &&
			!picture.natural;
		let size = { widthPt: picture.widthPt, heightPt: picture.heightPt };
		if (!unchanged) {
			const natural = picture.natural
				? size
				: ((await host.pictureSize?.(picture.partName, picture.contentType)) ?? size);
			size = pictureWatermarkSizePt(natural, chosen, host.textArea?.() ?? DEFAULT_AREA);
		}
		return {
			kind: 'picture',
			partName: picture.partName,
			contentType: picture.contentType,
			...size,
			scale: chosen,
			washout: washout.input.checked,
			...(picture.title ? { title: picture.title } : {}),
		};
	};
	const submit = async () => {
		if (!host.canEdit() || !validate()) return;
		const mode = setting.value;
		const spec: AnyWatermark | undefined =
			mode === 'picture'
				? await pictureSpec()
				: mode === 'text'
					? {
							text: text.value.trim(),
							color: color.value.toLowerCase(),
							semitransparent: semi.input.checked,
							layout: layout.value as 'diagonal' | 'horizontal',
							...(font.value.trim() ? { fontFamily: font.value.trim() } : {}),
						}
					: undefined;
		host.apply(spec);
		hide();
	};
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', () => void submit());
	onDialogDismiss(element, hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			void submit();
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
			const textSpec = current && !isPictureWatermark(current) ? current : undefined;
			initial = isPictureWatermark(current) ? current : undefined;
			picture = initial && { ...initial };
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			setting.value = initial ? 'picture' : textSpec ? 'text' : 'none';
			pictureName.textContent = initial?.title ?? '';
			scale.value = typeof initial?.scale === 'number' ? String(initial.scale) : 'auto';
			washout.input.checked = initial?.washout ?? true;
			text.value = textSpec?.text ?? '';
			font.value = textSpec?.fontFamily ?? 'Calibri';
			color.value = textSpec?.color ?? DEFAULT_COLOR;
			layout.value = textSpec?.layout ?? 'diagonal';
			semi.input.checked = textSpec?.semitransparent ?? true;
			validate();
			setDialogOpen(element, true);
			setting.focus();
		},
	};
}
