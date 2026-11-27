/**
 * Picture Format > Picture Layout: choose one of the built-in picture SmartArt layouts in the
 * shared SmartArt picker, then replace the selected pictures on their slide with that diagram.
 *
 * Loaded on demand from the picture ribbon, together with the built-in layout library (a
 * separate chunk), so a viewer that never opens it pays for neither. The conversion itself is
 * `pictureLayoutSmartArt` in `ooxml-core/pptx/smartart-layouts`.
 *
 * @module web-components/picture-layout-dialog
 */
import { isDirectPicture } from 'ooxml-core/pptx';
import type { PptxElement, PptxImageLikeElement, PptxSlide } from 'ooxml-core/pptx';

import { defineSmartArtPicker, type OfficeUiSmartArtPicker } from '../../smartart/smartart-picker';
import { canInteractWithElement } from '../render/element-locks';
import { generateElementId } from '../render/element-clipboard';
import { smartArtLayoutNameKey } from '../render/ribbon-galleries/smartart-layouts-gallery';
import { preloadSmartArtBuiltinLayouts } from '../render/smartart-builtin-layouts';
import { pictureDialog } from './picture-dialog-frame';
import type { PictureDialogContext } from './picture-ribbon';
import { canEditPicture } from './picture-ribbon-state';

/** The slide holding `id` among its top-level elements, with that slide's index. */
function slideOf(slides: readonly PptxSlide[], id: string) {
	const index = slides.findIndex((slide) => slide.elements.some((element) => element.id === id));
	return index < 0 ? undefined : { index, slide: slides[index]! };
}

/** The selected pictures on the picture's slide that may be converted. */
export function pictureLayoutTargets(
	slide: PptxSlide,
	primary: PptxImageLikeElement,
	selectedIds: readonly string[] | undefined,
): PptxImageLikeElement[] {
	const ids = new Set(selectedIds?.length ? selectedIds : [primary.id]);
	ids.add(primary.id);
	return slide.elements.filter(
		(element): element is PptxImageLikeElement =>
			ids.has(element.id) &&
			isDirectPicture(element) &&
			canEditPicture(element) &&
			canInteractWithElement(element, 'changeShapeType'),
	);
}

/** `elements` with the converted pictures removed and the diagram at the topmost one's place. */
export function replacePicturesWithDiagram(
	elements: readonly PptxElement[],
	pictures: readonly PptxImageLikeElement[],
	diagram: PptxElement,
): PptxElement[] {
	const ids = new Set(pictures.map((picture) => picture.id));
	const last = elements.reduce((top, element, index) => (ids.has(element.id) ? index : top), -1);
	return elements.flatMap((element, index) =>
		index === last ? [diagram] : ids.has(element.id) ? [] : [element],
	);
}

export async function choosePictureLayout(context: PictureDialogContext): Promise<void> {
	const { host, text, signal, picture } = context;
	const doc = host.ownerDocument;
	defineSmartArtPicker(doc.defaultView!.customElements);
	const picker = doc.createElement('office-ui-smartart-picker') as OfficeUiSmartArtPicker;
	picker.itemsEditor = false;
	const title = text('pptx.pictureFormat.layout', 'Picture Layout');
	const dialog = pictureDialog(context, title, () => void apply(), {
		apply: text('pptx.common.ok', 'OK'),
	});
	dialog.body.prepend(picker);
	dialog.status.textContent = text('pptx.pictureFormat.processing', 'Processing picture…');
	dialog.show();
	const library = await preloadSmartArtBuiltinLayouts();
	if (signal.aborted) return;
	if (!library) {
		dialog.fail();
		return;
	}
	dialog.status.textContent = '';
	picker.categories = [{ id: 'picture', label: title }];
	picker.layouts = library.listPictureLayouts().map((entry) => ({
		id: entry.id,
		category: 'picture',
		label: text(smartArtLayoutNameKey(entry.title), entry.title),
		family: entry.category === 'picture' ? 'list' : entry.category,
	}));
	picker.labels = {
		categories: title,
		layouts: text('pptx.insertSmartArt.layouts', 'SmartArt layouts'),
	};
	const ready = () => dialog.sync(Boolean(picker.value.layout));
	picker.addEventListener('office-smartart-change', ready);
	picker.addEventListener('office-smartart-activate', () => void apply());
	ready();

	async function apply(): Promise<void> {
		const layoutId = picker.value.layout;
		const state = context.state();
		const slides = state?.slides;
		const found = slides && slideOf(slides, picture.id);
		if (!layoutId || !library || !state?.editable || !state.updateSlides || !found) {
			dialog.fail();
			return;
		}
		dialog.sync(false);
		const targets = pictureLayoutTargets(found.slide, picture, state.selectedIds);
		const imageUrls = new Map<string, string>();
		for (const target of targets) {
			if (target.imageData || !target.imagePath) continue;
			const url = await state.loadImage?.(target.imagePath);
			if (url) imageUrls.set(target.id, url);
		}
		if (signal.aborted) return;
		const diagram = await library.pictureLayoutSmartArt(targets, layoutId, {
			id: generateElementId(),
			imageUrls,
		});
		if (signal.aborted) return;
		if (typeof diagram === 'string') {
			dialog.fail();
			return;
		}
		const slideId = found.slide.id;
		state.updateSlides((current) =>
			current.map((slide) =>
				slide.id === slideId
					? { ...slide, elements: replacePicturesWithDiagram(slide.elements, targets, diagram) }
					: slide,
			),
		);
		dialog.close();
	}
}
