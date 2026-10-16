/**
 * The contract every gallery module implements, plus the placeholder a
 * gallery uses before its catalogue lands.
 *
 * @module render/ribbon-galleries/gallery-module
 */
import type {
	RibbonGalleryApplyResult,
	RibbonGalleryContext,
	RibbonGalleryDescriptor,
	RibbonGalleryId,
} from './gallery-types';

export interface RibbonGalleryModule {
	build(ctx: RibbonGalleryContext): RibbonGalleryDescriptor;
	/** `action` is the context menu entry chosen on the tile, when one was. */
	apply(
		itemId: string,
		ctx: RibbonGalleryContext,
		action?: string,
	): RibbonGalleryApplyResult | null;
}

/** A gallery with no entries (disabled trigger). */
export function emptyGalleryModule(
	id: RibbonGalleryId,
	labelKey: string,
	label: string,
): RibbonGalleryModule {
	return {
		build: () => ({ id, labelKey, label, sections: [], disabled: true }),
		apply: () => null,
	};
}
