// Compatibility exports: document operations live in ooxml-core.
export {
	resolveTableCellImageUrls,
	resolveTableStyleImageUrls,
	resolveTextFillBlipUrls,
} from 'ooxml-core/pptx/editor/loader/lazy-image-resolution';
export type { GetImageData } from 'ooxml-core/pptx/editor/loader/lazy-image-resolution';
