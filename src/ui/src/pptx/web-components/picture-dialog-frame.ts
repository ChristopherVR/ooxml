import type { PptxImageLikeElement } from 'ooxml-core/pptx';
import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import type { PictureDialogContext } from './picture-ribbon';
import { canEditPicture } from './picture-ribbon-state';
import { localPictureBlob } from './picture-worker-client';

export function pictureDialog(
	context: PictureDialogContext,
	title: string,
	apply: () => void,
	options: { apply?: string; cancel?: string } = {},
) {
	const { host, text, signal } = context;
	const doc = host.ownerDocument;
	const body = doc.createElement('div');
	body.style.cssText = 'display:grid;gap:12px';
	const status = doc.createElement('div');
	status.setAttribute('role', 'status');
	body.append(status);
	const frame = createInsertDialogFrame(host, body, apply, context.cancel);
	frame.dialog.dataset.pictureDialog = '';
	signal.addEventListener(
		'abort',
		() => {
			frame.dialog.close();
			frame.dialog.remove();
		},
		{ once: true },
	);
	const sync = (ready: boolean, action = options.apply ?? text('pptx.common.apply', 'Apply')) =>
		frame.sync(
			title,
			text('pptx.common.close', 'Close'),
			options.cancel ?? text('pptx.common.cancel', 'Cancel'),
			action,
			ready,
		);
	return {
		body,
		status,
		sync,
		show() {
			sync(false);
			frame.dialog.show();
		},
		fail() {
			status.textContent = text(
				'pptx.pictureFormat.failed',
				'This picture could not be processed.',
			);
			sync(false);
		},
		close: context.cancel,
	};
}

export function pictureField(control: HTMLElement, label: string): HTMLLabelElement {
	const field = control.ownerDocument.createElement('label');
	field.style.cssText = 'display:flex;align-items:center;gap:8px';
	if (control.matches('office-ui-select, office-ui-checkbox'))
		control.setAttribute('aria-label', label);
	if (control.matches('office-ui-checkbox')) field.append(control, label);
	else field.append(label, control);
	return field;
}

export function commitPicture(
	context: PictureDialogContext,
	patch: Partial<PptxImageLikeElement>,
): boolean {
	const state = context.state();
	if (
		context.signal.aborted ||
		!state?.editable ||
		state.element !== context.picture ||
		!canEditPicture(context.picture)
	)
		return false;
	state.patch(context.picture.id, patch);
	return true;
}

export async function pictureSource(
	context: PictureDialogContext,
	picture: PptxImageLikeElement,
	original = false,
): Promise<Blob> {
	const path = original ? picture.originalPicturePath : picture.imagePath;
	const available = original ? picture.originalPictureData : picture.imageData;
	const packagePart =
		path && !/^(?:[a-z][a-z\d+.-]*:|\/\/)/iu.test(path) && /\.(?:png|jpe?g)$/iu.test(path);
	const source = available || (packagePart && (await context.state()?.loadImage?.(path)));
	if (!source) throw new Error('Picture source is unavailable');
	return localPictureBlob(source, context.signal);
}
