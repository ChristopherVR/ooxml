import {
	SHAPE_GALLERY_CATEGORIES,
	shapeGalleryCategoryKey,
	shapeGalleryNameKey,
} from 'ooxml-core/geometry';

/**
 * English copy for the Shapes gallery of Insert and Home: the category headings and every shape
 * name come from the core catalogue (`ooxml-core/geometry`), so a preset added there is named here
 * without a second list. The locales translate the same keys.
 */
export const shapeGalleryTranslationsEn: Record<string, string> = {
	'pptx.shapes.category.recent': 'Recently Used Shapes',
	'pptx.shapes.name.freeformScribble': 'Freeform: Scribble',
	'pptx.home.slideTemplatesMenu': 'Slide Templates...',
	'pptx.shapes.freeformFromInsert': 'Draw freeforms from Insert > Shapes.',
	...Object.fromEntries(
		SHAPE_GALLERY_CATEGORIES.flatMap((category) => [
			[shapeGalleryCategoryKey(category.id), category.name],
			...category.shapes.map((shape) => [shapeGalleryNameKey(shape.preset), shape.name]),
		]),
	),
};
