import {
	pictureBorderPatch,
	removePictureBackgroundPatch,
	replacePicturePatch,
	resetPicturePatch,
} from 'ooxml-core/pptx';
import { defineSelect, type OfficeUiSelect } from '../../form/select';
import { STROKE_DASH_OPTIONS } from '../render/stroke-dash-options';
import { commitPicture, pictureDialog, pictureField, pictureSource } from './picture-dialog-frame';
import type { PictureDialogCommand, PictureDialogContext } from './picture-ribbon';
import { createPictureWorker, pictureDataUrl } from './picture-worker-client';
import { MAX_PICTURE_MARKS } from './picture-segmentation';

export function runPictureDialog(
	action: PictureDialogCommand,
	context: PictureDialogContext,
): void {
	switch (action) {
		case 'changePicture':
			changePicture(context);
			break;
		case 'resetPicture':
		case 'resetSize':
			void resetPicture(context, action === 'resetSize');
			break;
		case 'border':
			editPictureBorder(context);
			break;
		case 'compress':
			void import('./picture-compress-dialog').then(({ compressPictures }) => {
				if (!context.signal.aborted) compressPictures(context);
			}, context.cancel);
			break;
		case 'layout':
			void import('./picture-layout-dialog').then(({ choosePictureLayout }) => {
				if (!context.signal.aborted) void choosePictureLayout(context);
			}, context.cancel);
			break;
		case 'removeBackground':
			void import('./picture-removal-dialog').then(({ removePictureBackground }) => {
				if (!context.signal.aborted) removePictureBackground(context);
			}, context.cancel);
			break;
	}
}

function changePicture(context: PictureDialogContext): void {
	const { host, text, signal } = context;
	const input = host.ownerDocument.createElement('input');
	input.type = 'file';
	input.accept = 'image/png,image/jpeg,.png,.jpg,.jpeg';
	const dialog = pictureDialog(
		context,
		text('pptx.pictureFormat.changePicture', 'Change Picture'),
		() => {
			const file = input.files?.[0];
			if (!file) return;
			dialog.sync(false);
			input.disabled = true;
			void (async () => {
				const worker = createPictureWorker(signal);
				try {
					const { source } = await worker.open(file);
					const data = await pictureDataUrl(source);
					const picture = context.picture;
					const removal = picture.imageEffects?.backgroundRemoval;
					const patch = replacePicturePatch(data);
					if (picture.imageEffects) {
						patch.imageEffects = { ...picture.imageEffects };
						delete patch.imageEffects.artisticPrerenderedEffect;
						delete patch.imageEffects.prerenderedCorrections;
						delete patch.imageEffects.originalImageRelId;
					}
					if (removal) {
						const result = await worker.send({
							kind: 'remove',
							marks: {
								foreground: (removal.foregroundMarks ?? []).slice(0, MAX_PICTURE_MARKS),
								background: (removal.backgroundMarks ?? []).slice(0, MAX_PICTURE_MARKS),
							},
							preview: false,
						});
						if (result.kind !== 'removed') throw new Error('Cannot apply background removal');
						Object.assign(
							patch,
							removePictureBackgroundPatch(
								{ ...picture, imageEffects: patch.imageEffects },
								await pictureDataUrl(result.image),
								data,
								result.removal,
							),
						);
					}
					if (commitPicture(context, patch)) dialog.close();
					else dialog.fail();
				} finally {
					worker.close();
					input.disabled = false;
				}
			})().catch(dialog.fail);
		},
	);
	input.addEventListener('change', () => dialog.sync(Boolean(input.files?.length)));
	dialog.body.prepend(
		pictureField(input, text('pptx.pictureFormat.chooseRaster', 'Choose a PNG or JPEG')),
	);
	dialog.show();
}

async function resetPicture(context: PictureDialogContext, resize: boolean): Promise<void> {
	const { text, picture, signal } = context;
	const title = resize
		? text('pptx.pictureFormat.resetSize', 'Reset Picture & Size')
		: text('pptx.pictureFormat.resetPicture', 'Reset Picture');
	const dialog = pictureDialog(context, title, () => dialog.close());
	dialog.status.textContent = text('pptx.pictureFormat.processing', 'Processing picture…');
	dialog.show();
	try {
		const hasOriginal = Boolean(picture.imageEffects?.backgroundRemoval);
		const source = await pictureSource(context, picture, hasOriginal);
		signal.throwIfAborted();
		const worker = createPictureWorker(signal);
		try {
			const { source: validated, width, height } = await worker.open(source);
			const patch = resetPicturePatch(picture, {
				...(hasOriginal && { original: await pictureDataUrl(validated) }),
				...(resize && { size: { width, height } }),
			});
			if (commitPicture(context, patch)) dialog.close();
			else dialog.fail();
		} finally {
			worker.close();
		}
	} catch {
		dialog.fail();
	}
}

function editPictureBorder(context: PictureDialogContext): void {
	const { host, picture, text } = context;
	const doc = host.ownerDocument;
	const color = doc.createElement('input');
	color.type = 'color';
	const stroke = picture.shapeStyle?.strokeColor;
	color.value = stroke && /^#[\da-f]{6}$/iu.test(stroke) ? stroke : '#000000';
	const weight = doc.createElement('input');
	weight.type = 'number';
	weight.min = '0';
	weight.max = '75';
	weight.step = 'any';
	weight.required = true;
	weight.value = String(Math.round((picture.shapeStyle?.strokeWidth || 4 / 3) * 75) / 100);
	defineSelect(doc.defaultView!.customElements);
	const dash = doc.createElement('office-ui-select') as OfficeUiSelect;
	dash.options = STROKE_DASH_OPTIONS.filter(({ value }) => value !== 'custom').map((option) => ({
		value: option.value,
		label: text(option.i18nKey, option.label),
	}));
	dash.value = picture.shapeStyle?.strokeDash ?? 'solid';
	const dialog = pictureDialog(context, text('pptx.pictureFormat.border', 'Picture Border'), () => {
		const selected = STROKE_DASH_OPTIONS.find((option) => option.value === dash.value);
		if (!selected || !weight.validity.valid) return;
		if (
			commitPicture(
				context,
				pictureBorderPatch(picture, {
					color: color.value,
					width: (weight.valueAsNumber * 4) / 3,
					dash: selected.value,
				}),
			)
		)
			dialog.close();
		else dialog.fail();
	});
	weight.addEventListener('input', () => dialog.sync(weight.validity.valid));
	dialog.body.prepend(
		pictureField(color, text('pptx.pictureFormat.borderColor', 'Color')),
		pictureField(weight, text('pptx.pictureFormat.borderWeight', 'Weight (pt)')),
		pictureField(dash, text('pptx.strokeEffects.strokeDash', 'Dashes')),
	);
	dialog.show();
	dialog.sync(true);
}
