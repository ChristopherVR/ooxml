import type { PptxElement } from 'ooxml-core/pptx';
import { normalizeStShapeType } from 'ooxml-core/geometry';
import { cropShapeForPresetGeometry } from 'ooxml-core/pptx';
import { canInteractWithElement } from '../element-locks';
import { SHAPE_THUMB_VIEW_BOX, shapeGalleryItem, shapeGallerySections } from '../shape-gallery';
import type { RibbonGalleryModule } from './gallery-module';
import type { RibbonGallerySection } from './gallery-types';
import { isPictureElement } from './picture-adjust-gallery';

export function presetShapeSections(current?: string): RibbonGallerySection[] {
	const selected = normalizeStShapeType(current);
	return shapeGallerySections([]).map(({ id, heading, headingKey, items }) => ({
		id,
		title: heading,
		titleKey: headingKey,
		columns: 8,
		tileWidth: 26,
		tileHeight: 26,
		items: items
			.filter(({ action }) => action.kind === 'shape' || action.kind === 'actionButton')
			.map(({ id, labelKey, fallback, path }) => ({
				id,
				labelKey,
				label: fallback,
				previewSvg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${SHAPE_THUMB_VIEW_BOX}" width="26" height="26"><path d="${path}" fill="none" stroke="currentColor"/></svg>`,
				applied: id === selected,
			})),
	}));
}

function pictureTarget(element: PptxElement | null): element is PptxElement {
	return (
		isPictureElement(element) &&
		canInteractWithElement(element, 'changeShapeType') &&
		!(element.type !== 'media' && 'liveFeed' in element && element.liveFeed)
	);
}

export const CROP_TO_SHAPE_GALLERY: RibbonGalleryModule = {
	build({ element }) {
		const media = element?.type === 'media';
		return {
			id: 'cropToShape',
			labelKey: media ? 'pptx.mediaRibbon.changeShape' : 'pptx.mediaRibbon.cropToShape',
			label: media ? 'Change Shape' : 'Crop to Shape',
			disabled: !pictureTarget(element),
			sections: presetShapeSections(
				element?.type === 'media' || element?.type === 'picture' || element?.type === 'image'
					? (element.shapeType ?? 'rect')
					: undefined,
			),
		};
	},
	apply(shapeType, { element }) {
		const item = shapeGalleryItem(shapeType);
		if (
			!pictureTarget(element) ||
			(item?.action.kind !== 'shape' && item?.action.kind !== 'actionButton')
		)
			return null;
		const patch: Partial<PptxElement> = {
			shapeType,
			shapeAdjustments: undefined,
			pathData: undefined,
			customGeometryPaths: undefined,
			customGeometryRawData: undefined,
			...(element.type !== 'media' && { cropShape: cropShapeForPresetGeometry(shapeType) }),
		};
		return { kind: 'element', elementId: element.id, patch };
	},
};
