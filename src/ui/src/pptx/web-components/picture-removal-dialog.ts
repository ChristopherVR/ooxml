import { removePictureBackgroundPatch } from 'ooxml-core/pptx';
import type { PptxBackgroundRemovalMark } from 'ooxml-core/pptx';
import { getImageBackgroundRemoval } from '../render/image-background-removal';
import { commitPicture, pictureDialog, pictureSource } from './picture-dialog-frame';
import type { PictureDialogContext } from './picture-ribbon';
import { MAX_PICTURE_MARKS } from './picture-segmentation';
import { createPictureWorker, pictureDataUrl } from './picture-worker-client';

const limit = (value: number) => Math.max(0, Math.min(1, value));
type Point = readonly [number, number];

export function removePictureBackground(context: PictureDialogContext): void {
	const { host, text, picture, signal } = context;
	const doc = host.ownerDocument;
	const pictureArea = doc.createElement('button');
	pictureArea.type = 'button';
	pictureArea.style.cssText =
		'position:relative;display:block;max-width:100%;width:max-content;padding:0;border:0;touch-action:none;cursor:crosshair;background:var(--office-surface-muted, #dfdfdf)';
	pictureArea.setAttribute(
		'aria-label',
		text('pptx.pictureFormat.picturePreview', 'Picture preview'),
	);
	pictureArea.hidden = true;
	const image = doc.createElement('img');
	image.alt = '';
	image.draggable = false;
	image.style.cssText = 'display:block;max-width:min(100%, calc(100vw - 80px));max-height:55vh';
	const marquee = doc.createElement('span');
	marquee.dataset.pptxRemovalMarquee = '';
	marquee.setAttribute('aria-hidden', 'true');
	marquee.style.cssText =
		'position:absolute;border:1px dashed var(--office-ring,#0078d4);pointer-events:none';
	const cursor = doc.createElement('span');
	cursor.style.cssText =
		'position:absolute;width:12px;height:12px;border:2px solid white;outline:1px solid black;border-radius:50%;transform:translate(-50%,-50%);pointer-events:none';
	cursor.hidden = true;
	pictureArea.append(image, marquee, cursor);
	const marks = {
		foreground: (picture.imageEffects?.backgroundRemoval?.foregroundMarks ?? []).slice(
			0,
			MAX_PICTURE_MARKS,
		),
		background: (picture.imageEffects?.backgroundRemoval?.backgroundMarks ?? []).slice(
			0,
			MAX_PICTURE_MARKS,
		),
	};
	let mode: keyof typeof marks = 'foreground';
	let point: Point = [0.5, 0.5];
	let start: Point | undefined;
	let url: string | undefined;
	let busy = true;
	let worker: ReturnType<typeof createPictureWorker> | undefined;
	let original: Blob;
	const options = doc.createElement('div');
	options.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap';
	const foreground = doc.createElement('button');
	const background = doc.createElement('button');
	foreground.type = background.type = 'button';
	foreground.textContent = text('pptx.pictureFormat.markKeep', 'Mark Areas to Keep');
	background.textContent = text('pptx.pictureFormat.markRemove', 'Mark Areas to Remove');
	options.append(foreground, background);
	const choose = (value: typeof mode) => {
		mode = value;
		foreground.setAttribute('aria-pressed', String(value === 'foreground'));
		background.setAttribute('aria-pressed', String(value === 'background'));
	};
	choose(mode);
	foreground.addEventListener('click', () => choose('foreground'));
	background.addEventListener('click', () => choose('background'));
	const dialog = pictureDialog(
		context,
		text('pptx.pictureFormat.removeBackground', 'Remove Background'),
		() => {
			if (!busy) void render(false).catch(dialog.fail);
		},
		{
			apply: text('pptx.pictureFormat.keepChanges', 'Keep Changes'),
			cancel: text('pptx.pictureFormat.discard', 'Discard'),
		},
	);
	const help = doc.createElement('p');
	help.id = 'pptx-picture-mark-instructions';
	help.textContent = text(
		'pptx.pictureFormat.markInstructions',
		'Drag to mark an area, or move with arrow keys and press Space to mark a point.',
	);
	pictureArea.setAttribute('aria-describedby', help.id);
	dialog.body.prepend(options, help, pictureArea);
	function mark(from: Point, to: Point): void {
		if (busy) return;
		if (marks.foreground.length + marks.background.length >= MAX_PICTURE_MARKS) {
			dialog.status.textContent = text(
				'pptx.pictureFormat.markLimit',
				'The picture has reached the mark limit.',
			);
			return;
		}
		const stroke: PptxBackgroundRemovalMark = { x1: from[0], y1: from[1], x2: to[0], y2: to[1] };
		marks[mode].push(stroke);
		void render(true).catch(dialog.fail);
	}
	function pointer(event: PointerEvent): Point {
		const bounds = pictureArea.getBoundingClientRect();
		return [
			limit((event.clientX - bounds.left) / bounds.width),
			limit((event.clientY - bounds.top) / bounds.height),
		];
	}
	pictureArea.addEventListener('pointerdown', (event) => {
		if (!busy && event.isPrimary && event.button === 0) {
			start = pointer(event);
			pictureArea.setPointerCapture(event.pointerId);
		}
	});
	pictureArea.addEventListener('pointerup', (event) => {
		if (start) mark(start, pointer(event));
		start = undefined;
	});
	pictureArea.addEventListener('pointercancel', () => {
		start = undefined;
	});
	pictureArea.addEventListener('focus', () => moveCursor(point));
	pictureArea.addEventListener('blur', () => {
		cursor.hidden = true;
	});
	pictureArea.addEventListener('keydown', (event) => {
		const step = event.shiftKey ? 0.1 : 0.02;
		switch (event.key) {
			case 'ArrowLeft':
				moveCursor([point[0] - step, point[1]]);
				break;
			case 'ArrowRight':
				moveCursor([point[0] + step, point[1]]);
				break;
			case 'ArrowUp':
				moveCursor([point[0], point[1] - step]);
				break;
			case 'ArrowDown':
				moveCursor([point[0], point[1] + step]);
				break;
			case ' ':
			case 'Enter':
				mark(point, point);
				break;
			default:
				return;
		}
		event.preventDefault();
	});
	function moveCursor(next: Point): void {
		point = [limit(next[0]), limit(next[1])];
		cursor.style.left = `${point[0] * 100}%`;
		cursor.style.top = `${point[1] * 100}%`;
		cursor.hidden = false;
	}
	signal.addEventListener(
		'abort',
		() => {
			worker?.close();
			if (url) URL.revokeObjectURL(url);
		},
		{ once: true },
	);
	dialog.show();
	dialog.status.textContent = text('pptx.pictureFormat.processing', 'Processing picture…');
	void (async () => {
		const blob = await pictureSource(
			context,
			picture,
			Boolean(picture.imageEffects?.backgroundRemoval),
		);
		signal.throwIfAborted();
		worker = createPictureWorker(signal);
		original = (await worker.open(blob)).source;
		await render(true);
	})().catch(dialog.fail);

	async function render(preview: boolean): Promise<void> {
		if (!worker) return;
		busy = true;
		dialog.sync(false);
		const result = await worker.send({ kind: 'remove', marks, preview });
		signal.throwIfAborted();
		if (result.kind !== 'removed') throw new Error('Background removal failed');
		if (preview) {
			if (url) URL.revokeObjectURL(url);
			url = URL.createObjectURL(result.image);
			image.src = url;
			const descriptor = getImageBackgroundRemoval({
				...picture,
				imageEffects: { backgroundRemoval: result.removal },
			});
			if (descriptor) {
				const { top, right, bottom, left } = descriptor.retained;
				marquee.style.inset = `${top * 100}% ${(1 - right) * 100}% ${(1 - bottom) * 100}% ${left * 100}%`;
			}
			pictureArea.hidden = false;
			dialog.status.textContent = '';
			busy = false;
			dialog.sync(true);
		} else {
			const [data, source] = await Promise.all([
				pictureDataUrl(result.image),
				pictureDataUrl(original),
			]);
			if (
				commitPicture(context, removePictureBackgroundPatch(picture, data, source, result.removal))
			)
				dialog.close();
			else dialog.fail();
		}
	}
}
