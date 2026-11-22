import { createInsertDialogFrame } from '../dialog/insert-dialog-frame';
import { captureDisplayFrame } from './display-capture';

export interface ScreenshotLabels {
	heading: string;
	close: string;
	cancel: string;
	insert: string;
	x: string;
	y: string;
	width: string;
	height: string;
}

export async function captureScreenshot(
	host: HTMLElement | ShadowRoot,
	clipping?: ScreenshotLabels,
) {
	const image = await captureDisplayFrame(host.ownerDocument);
	if (!image || !host.isConnected) return undefined;
	return clipping ? pickClipping(host, image, clipping) : image;
}

export async function captureScreenshotFile(
	host: HTMLElement | ShadowRoot,
	name: string,
	clipping?: ScreenshotLabels,
): Promise<File | undefined> {
	const screenshot = await captureScreenshot(host, clipping);
	if (!screenshot) return undefined;
	const blob = await new Promise<Blob | null>((resolve) => screenshot.toBlob(resolve, 'image/png'));
	return blob ? new File([blob], name + '.png', { type: 'image/png' }) : undefined;
}

export interface DisplayRegion {
	x: number;
	y: number;
	width: number;
	height: number;
}

function pickClipping(
	host: HTMLElement | ShadowRoot,
	image: HTMLCanvasElement,
	labels: ScreenshotLabels,
) {
	return new Promise<HTMLCanvasElement | undefined>((resolve) => {
		let selected: DisplayRegion | undefined;
		const picker = createDisplayRegionPicker(image, labels, (region) => {
			selected = region;
			sync();
		});
		const doc = host.ownerDocument;
		const frame = createInsertDialogFrame(
			host,
			picker.element,
			() => {
				if (!selected) return;
				const { x, y, width, height } = selected;
				const crop = doc.createElement('canvas');
				crop.width = width;
				crop.height = height;
				const output = crop.getContext('2d');
				output?.drawImage(image, x, y, width, height, 0, 0, width, height);
				resolve(output ? crop : undefined);
				frame.dialog.close();
			},
			() => {
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		function sync() {
			frame.sync(labels.heading, labels.close, labels.cancel, labels.insert, !!selected);
		}
		picker.update();
		frame.dialog.show();
	});
}

export function createDisplayRegionPicker(
	image: HTMLCanvasElement,
	labels: Pick<ScreenshotLabels, 'heading' | 'x' | 'y' | 'width' | 'height'>,
	onChange: (region: DisplayRegion | undefined) => void,
	minimum = 1,
) {
	const doc = image.ownerDocument;
	const body = doc.createElement('div');
	const preview = doc.createElement('canvas');
	preview.width = image.width;
	preview.height = image.height;
	preview.style.cssText = 'display:block;width:100%;height:auto;touch-action:none;cursor:crosshair';
	preview.setAttribute('aria-label', labels.heading);
	const context = preview.getContext('2d');
	if (!context) throw new Error('Could not render screen preview');
	const fields = doc.createElement('div');
	fields.style.cssText = 'display:flex;flex-wrap:wrap;gap:8px;margin-block:8px';
	function field(label: string, initial: number, min: number, max: number): HTMLInputElement {
		const caption = doc.createElement('label');
		const control = doc.createElement('input');
		control.type = 'number';
		control.required = true;
		control.value = String(initial);
		control.min = String(min);
		control.max = String(max);
		control.style.width = '8ch';
		caption.append(label, ' ', control);
		fields.append(caption);
		return control;
	}
	const x = field(labels.x, 0, 0, image.width - 1);
	const y = field(labels.y, 0, 0, image.height - 1);
	const width = field(labels.width, image.width, Math.min(minimum, image.width), image.width);
	const height = field(labels.height, image.height, Math.min(minimum, image.height), image.height);
	const inputs = [x, y, width, height];
	body.append(preview, fields);
	const sync = () => {
		const box = {
			x: x.valueAsNumber,
			y: y.valueAsNumber,
			width: width.valueAsNumber,
			height: height.valueAsNumber,
		};
		context.drawImage(image, 0, 0);
		const valid =
			inputs.every((input) => input.validity.valid) &&
			box.x + box.width <= image.width &&
			box.y + box.height <= image.height;
		if (valid && (box.width !== image.width || box.height !== image.height)) {
			context.fillStyle = 'rgba(0,0,0,.55)';
			context.fillRect(0, 0, image.width, image.height);
			context.drawImage(
				image,
				box.x,
				box.y,
				box.width,
				box.height,
				box.x,
				box.y,
				box.width,
				box.height,
			);
		}
		onChange(valid ? box : undefined);
	};
	const point = (event: PointerEvent) => {
		const rect = preview.getBoundingClientRect();
		return {
			x: Math.round(
				Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width)) * image.width,
			),
			y: Math.round(
				Math.max(0, Math.min(1, (event.clientY - rect.top) / rect.height)) * image.height,
			),
		};
	};
	let origin: { x: number; y: number } | undefined;
	preview.addEventListener('pointerdown', (event) => {
		if (!event.isPrimary || event.button !== 0) return;
		origin = point(event);
		preview.setPointerCapture(event.pointerId);
	});
	preview.addEventListener('pointermove', (event) => {
		if (!event.isPrimary || !origin) return;
		const cursor = point(event);
		x.value = String(Math.min(image.width - 1, origin.x, cursor.x));
		y.value = String(Math.min(image.height - 1, origin.y, cursor.y));
		width.value = String(Math.max(1, Math.abs(origin.x - cursor.x)));
		height.value = String(Math.max(1, Math.abs(origin.y - cursor.y)));
		sync();
	});
	preview.addEventListener('lostpointercapture', () => {
		origin = undefined;
	});
	for (const input of inputs) input.addEventListener('input', sync);
	return { element: body, update: sync };
}
