import { clearPictureCrop, isDirectPicture, replacePicturePatch } from 'ooxml-core/pptx';
import type { PptxElement, PptxImageLikeElement, PptxSlide } from 'ooxml-core/pptx';
import { canInteractWithElement } from '../render/element-locks';
import { commitPicture, pictureDialog, pictureField, pictureSource } from './picture-dialog-frame';
import type { PictureDialogContext } from './picture-ribbon';
import { canEditPicture, canProcessPictureRaster } from './picture-ribbon-state';
import { createPictureWorker, pictureDataUrl } from './picture-worker-client';

const MAX_BATCH_PICTURES = 200;
const MAX_BATCH_OUTPUT_BYTES = 64 * 1024 * 1024;

function* picturesIn(elements: readonly PptxElement[], depth = 0): Generator<PptxImageLikeElement> {
	for (const element of elements) {
		if (isDirectPicture(element)) yield element;
		else if (element.type === 'group' && depth < 32) yield* picturesIn(element.children, depth + 1);
	}
}

function applyCompression(
	slides: readonly PptxSlide[],
	patches: ReadonlyMap<PptxImageLikeElement, Partial<PptxImageLikeElement>>,
): PptxSlide[] {
	const edit = (list: PptxElement[], depth = 0): PptxElement[] => {
		let changed = false;
		const result = list.map((element): PptxElement => {
			if (element.type === 'group' && depth < 32) {
				const children = edit(element.children, depth + 1);
				if (children !== element.children) {
					changed = true;
					return { ...element, children };
				}
			}
			if (isDirectPicture(element)) {
				const patch = patches.get(element);
				if (patch) {
					changed = true;
					// Keep the discriminant and all non-picture fields of the source.
					return Object.assign({ ...element }, patch);
				}
			}
			return element;
		});
		return changed ? result : list;
	};
	return slides.map((slide) => {
		const elements = edit(slide.elements);
		return elements === slide.elements ? slide : { ...slide, elements };
	});
}

export function compressPictures(context: PictureDialogContext): void {
	const { host, picture, text, signal } = context;
	const doc = host.ownerDocument;
	const ppi = doc.createElement('select');
	for (const resolution of [330, 220, 150, 96]) {
		const option = doc.createElement('option');
		option.value = option.textContent = String(resolution);
		ppi.append(option);
	}
	ppi.value = '150';
	const all = doc.createElement('input');
	all.type = 'checkbox';
	const crop = doc.createElement('input');
	crop.type = 'checkbox';
	const form = doc.createElement('div');
	form.style.cssText = 'display:grid;gap:8px';
	form.append(
		pictureField(ppi, text('pptx.pictureFormat.resolution', 'Resolution (pixels per inch)')),
	);
	if (context.state()?.slides && context.state()?.updateSlides) {
		form.append(
			pictureField(
				all,
				text('pptx.pictureFormat.allPictures', 'All pictures in this presentation'),
			),
		);
	}
	form.append(
		pictureField(crop, text('pptx.pictureFormat.deleteCrop', 'Delete cropped areas of pictures')),
	);
	let done = false;
	const dialog = pictureDialog(
		context,
		text('pptx.pictureFormat.compress', 'Compress Pictures'),
		() => {
			if (done) {
				dialog.close();
				return;
			}
			dialog.sync(false);
			ppi.disabled = all.disabled = crop.disabled = true;
			dialog.status.textContent = text('pptx.pictureFormat.processing', 'Processing picture…');
			void apply().catch(dialog.fail);
		},
	);
	dialog.body.prepend(form);
	dialog.show();
	dialog.sync(true);

	async function apply(): Promise<void> {
		const snapshot = context.state();
		const slides = all.checked ? snapshot?.slides : undefined;
		const candidates = slides
			? slides.flatMap((slide) => [...picturesIn(slide.elements)])
			: [picture];
		const patches = new Map<PptxImageLikeElement, Partial<PptxImageLikeElement>>();
		const worker = createPictureWorker(signal);
		let saved = 0;
		let updatedBytes = 0;
		let skipped = 0;
		try {
			for (let index = 0; index < candidates.length; index++) {
				signal.throwIfAborted();
				const entry = candidates[index];
				if (
					!entry ||
					!canEditPicture(entry) ||
					!canProcessPictureRaster(entry) ||
					entry.imageEffects?.backgroundRemoval ||
					entry.imageEffects?.prerenderedCorrections ||
					entry.imageEffects?.artisticPrerenderedEffect ||
					index >= MAX_BATCH_PICTURES ||
					updatedBytes >= MAX_BATCH_OUTPUT_BYTES
				) {
					skipped++;
					continue;
				}
				try {
					const source = await pictureSource(context, entry);
					await worker.open(source);
					const result = await worker.send({
						kind: 'compress',
						options: {
							width: entry.width,
							height: entry.height,
							ppi: Number(ppi.value),
							crop: [
								entry.cropLeft ?? 0,
								entry.cropTop ?? 0,
								entry.cropRight ?? 0,
								entry.cropBottom ?? 0,
							],
							deleteCrop:
								crop.checked &&
								canInteractWithElement(entry, 'crop') &&
								entry.tileScaleX === undefined &&
								entry.tileScaleY === undefined,
						},
					});
					if (result.kind !== 'compressed') throw new Error('Picture compression failed');
					if (!result.deletedCrop && result.image.size >= source.size) continue;
					if (updatedBytes + result.image.size > MAX_BATCH_OUTPUT_BYTES) {
						skipped++;
						continue;
					}
					const data = await pictureDataUrl(result.image);
					patches.set(entry, {
						...replacePicturePatch(data),
						...(result.deletedCrop && clearPictureCrop()),
					});
					updatedBytes += result.image.size;
					saved += source.size - result.image.size;
				} catch {
					skipped++;
				}
			}
		} finally {
			worker.close();
		}
		signal.throwIfAborted();
		const current = context.state();
		if (!current?.editable || current.element !== picture) {
			dialog.fail();
			return;
		}
		if (slides) {
			if (!current.updateSlides || current.slides !== slides) {
				dialog.fail();
				return;
			}
			if (patches.size) current.updateSlides((latest) => applyCompression(latest, patches));
		} else {
			const patch = patches.get(picture);
			if (patch && !commitPicture(context, patch)) {
				dialog.fail();
				return;
			}
		}
		done = true;
		form.hidden = true;
		const format = new Intl.NumberFormat(doc.documentElement.lang || undefined);
		dialog.status.textContent = [
			text('pptx.pictureFormat.bytesSaved', 'Picture bytes saved'),
			format.format(saved),
			'·',
			text('pptx.pictureFormat.picturesSkipped', 'Pictures skipped'),
			format.format(skipped),
		].join(' ');
		dialog.sync(true, text('pptx.common.ok', 'OK'));
	}
}
