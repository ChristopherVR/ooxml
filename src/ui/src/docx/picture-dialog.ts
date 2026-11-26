import type { EditorView } from 'prosemirror-view';
import {
	isValidPictureCrop,
	pictureCropOf,
	setPictureCrop,
	uncroppedPictureSize,
} from 'ooxml-core/docx';
import type { CropInsets } from 'ooxml-core/geometry';
import { focusView } from './focus-view';
import { fieldset, labelled, numberInput, row } from './dialog-fields';
import { localizeElement, translate, type EditorLocale } from './localization';
import {
	isPictureNode,
	pictureOf,
	VML_CROP_REASON,
	VML_EDIT_REASON,
} from './picture-format-commands';
import { schema } from './schema';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

export interface PictureDialog {
	element: HTMLElement;
	/** Opens on the picture at `pos`; `focus: 'crop'` starts in the Crop fields (Picture Format > Crop). */
	open(pos: number, focus?: 'crop'): void;
	close(): void;
	setLocale(locale: EditorLocale): void;
	readonly isOpen: boolean;
}

const CROP_SIDES: Array<[keyof CropInsets, string]> = [
	['cropLeft', 'Left'],
	['cropRight', 'Right'],
	['cropTop', 'Top'],
	['cropBottom', 'Bottom'],
];
const CROP_ERROR = 'The crop must leave part of the picture showing.';
const percent = (fraction: number) => String(+(fraction * 100).toFixed(2));

/** Word's Format Picture essentials: alt text, size (px) with an aspect-ratio lock, and crop. */
export function createPictureDialog(getView: () => EditorView | undefined): PictureDialog {
	const element = createDialogShell('Format picture', 'dve-picture-dialog');
	const alt = document.createElement('textarea');
	alt.rows = 3;
	const width = numberInput(16, 20000, 1);
	const height = numberInput(16, 20000, 1);
	const lock = document.createElement('input');
	lock.type = 'checkbox';
	lock.checked = true;
	const lockLabel = document.createElement('label');
	lockLabel.className = 'dve-dialog-check';
	const lockText = document.createElement('span');
	lockText.textContent = 'Lock aspect ratio';
	lockLabel.append(lock, lockText);
	const sizes = document.createElement('div');
	sizes.className = 'dve-dialog-row';
	sizes.append(labelled('Width', width), labelled('Height', height));
	const crops = new Map(CROP_SIDES.map(([side]) => [side, numberInput(-100, 99.99, 0.01)]));
	const cropField = (side: keyof CropInsets, label: string) => labelled(label, crops.get(side)!);
	const cropSet = fieldset(
		'Crop (%)',
		row(cropField('cropLeft', 'Left'), cropField('cropRight', 'Right')),
		row(cropField('cropTop', 'Top'), cropField('cropBottom', 'Bottom')),
	);
	cropSet.classList.add('dve-picture-crop');
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = document.createElement('button');
	cancel.type = 'button';
	cancel.textContent = 'Cancel';
	const apply = document.createElement('button');
	apply.type = 'button';
	apply.className = 'dve-dialog-primary';
	apply.textContent = 'OK';
	const actions = dialogActions(cancel, apply);
	element.append(labelled('Alt text', alt), sizes, lockLabel, cropSet, message, actions);

	let locale: EditorLocale = 'en';
	let position = -1;
	let ratio = 1;
	let whole = { widthPx: 1, heightPx: 1 };
	let initialCrop: CropInsets = { cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0 };
	const readCrop = (): CropInsets | undefined => {
		const crop = Object.fromEntries(
			CROP_SIDES.map(([side]) => [side, Number(crops.get(side)!.value || 0) / 100]),
		) as unknown as CropInsets;
		return isValidPictureCrop(crop) ? crop : undefined;
	};
	const cropChanged = (crop: CropInsets) =>
		CROP_SIDES.some(([side]) => Math.abs(crop[side] - initialCrop[side]) > 1e-6);
	const updateLegacyDescriptions = () => {
		if (cropSet.disabled) {
			cropSet.setAttribute('aria-description', translate(locale, VML_CROP_REASON));
			alt.setAttribute('aria-description', translate(locale, VML_EDIT_REASON));
		} else {
			cropSet.removeAttribute('aria-description');
			alt.removeAttribute('aria-description');
		}
	};
	const validate = () => {
		const valid = cropSet.disabled || Boolean(readCrop());
		message.textContent = cropSet.disabled ? VML_CROP_REASON : valid ? '' : CROP_ERROR;
		localizeElement(message, locale);
		apply.disabled = !valid;
		return valid;
	};
	width.addEventListener('input', () => {
		if (lock.checked && Number(width.value) > 0)
			height.value = String(Math.round(Number(width.value) * ratio));
	});
	height.addEventListener('input', () => {
		if (lock.checked && Number(height.value) > 0)
			width.value = String(Math.round(Number(height.value) / ratio));
	});
	// Cropping keeps the picture's scale, so the shown size follows the crop as Word's does.
	for (const input of crops.values())
		input.addEventListener('input', () => {
			if (cropSet.disabled) return;
			const crop = validate() && readCrop();
			if (!crop) return;
			const shownWidth = whole.widthPx * (1 - crop.cropLeft - crop.cropRight);
			const shownHeight = whole.heightPx * (1 - crop.cropTop - crop.cropBottom);
			width.value = String(Math.max(16, Math.round(shownWidth)));
			height.value = String(Math.max(16, Math.round(shownHeight)));
			ratio = shownHeight / Math.max(1, shownWidth);
		});
	const close = () => {
		setDialogOpen(element, false);
		focusView(getView());
	};
	const submit = () => {
		const view = getView();
		const node = view?.state.doc.nodeAt(position);
		if (!view || !node || node.type !== schema.nodes.image) return close();
		const crop = node.attrs.vmlFloat ? undefined : readCrop();
		if (!node.attrs.vmlFloat && !crop) return void validate();
		const cropped =
			crop && isPictureNode(node) && cropChanged(crop)
				? setPictureCrop(pictureOf(node), crop)
				: undefined;
		const size = (input: HTMLInputElement, current: number) =>
			node.attrs.vmlFloat && input.value === String(Math.round(current))
				? current
				: Math.max(16, Math.round(Number(input.value) || current));
		view.dispatch(
			view.state.tr.setNodeMarkup(position, undefined, {
				...node.attrs,
				...(node.attrs.vmlFloat ? {} : { altText: alt.value.trim() || null }),
				widthPx: size(width, node.attrs.widthPx),
				heightPx: size(height, node.attrs.heightPx),
				...(cropped ? { pictureFill: JSON.stringify(cropped.pictureFill) } : {}),
			}),
		);
		close();
	};
	cancel.addEventListener('click', close);
	apply.addEventListener('click', submit);
	onDialogDismiss(element, close);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});
	return {
		element,
		open(pos, focus) {
			const node = getView()?.state.doc.nodeAt(pos);
			if (!node || node.type !== schema.nodes.image) return;
			position = pos;
			ratio = Number(node.attrs.heightPx) / Math.max(1, Number(node.attrs.widthPx));
			alt.value = node.attrs.altText ?? '';
			width.value = String(Math.round(node.attrs.widthPx));
			height.value = String(Math.round(node.attrs.heightPx));
			// Only a real picture has a source rectangle to crop.
			cropSet.hidden = !isPictureNode(node);
			cropSet.disabled = alt.disabled = node.attrs.vmlFloat === true;
			updateLegacyDescriptions();
			const picture = pictureOf(node);
			initialCrop = pictureCropOf(picture);
			whole = uncroppedPictureSize(picture);
			for (const [side] of CROP_SIDES) crops.get(side)!.value = percent(initialCrop[side]);
			validate();
			setDialogOpen(element, true);
			const initial = cropSet.disabled
				? width
				: focus === 'crop' && !cropSet.hidden
					? crops.get('cropLeft')
					: alt;
			initial?.focus();
		},
		close,
		setLocale(next) {
			locale = next;
			localizeDialog(element, next);
			updateLegacyDescriptions();
		},
		get isOpen() {
			return isDialogOpen(element);
		},
	};
}
