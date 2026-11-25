import type { PptxElement, PptxImageLikeElement, PptxSlide } from 'ooxml-core/pptx';
import { canInteractWithElement } from '../render/element-locks';

export interface PictureRibbonState {
	element: PptxElement | null | undefined;
	editable: boolean;
	translate: (key: string) => string;
	patch: (id: string, patch: Partial<PptxImageLikeElement>) => void;
	altText?: (() => void) | undefined;
	selectionPane?: (() => void) | undefined;
	loadImage?: ((path: string) => Promise<string | undefined>) | undefined;
	slides?: readonly PptxSlide[] | undefined;
	updateSlides?: ((update: (slides: PptxSlide[]) => PptxSlide[]) => void) | undefined;
}

export type PictureStripGroup =
	| 'picture-adjust'
	| 'picture-styles'
	| 'picture-accessibility'
	| 'picture-arrange';

export function canEditPicture(element: PptxImageLikeElement): boolean {
	return canInteractWithElement(element, 'select') && !('liveFeed' in element && element.liveFeed);
}

export function canProcessPictureRaster(picture: PptxImageLikeElement): boolean {
	if (picture.svgData || picture.svgPath) return false;
	if (picture.imageDataEdited) return true;
	if (
		picture.imagePath &&
		(!/\.(?:png|jpe?g)$/iu.test(picture.imagePath) ||
			/^(?:[a-z][a-z\d+.-]*:|\/\/)/iu.test(picture.imagePath))
	)
		return false;
	return Boolean(
		picture.imagePath ||
		picture.imageData?.startsWith('blob:') ||
		/^data:image\/(?:png|jpe?g);base64,/iu.test(picture.imageData ?? ''),
	);
}
