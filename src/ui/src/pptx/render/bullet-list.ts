// Compatibility exports: document operations live in ooxml-core.
export {
	resolvePictureBullet,
	resolveParagraphBullet,
	bulletIndentPx,
	resolveParagraphIndent,
} from 'ooxml-core/pptx/editor/render/bullet-list';
export type {
	ParagraphBulletResult,
	PictureBulletMarker,
	ParagraphIndent,
	ParagraphIndentLayout,
} from 'ooxml-core/pptx/editor/render/bullet-list';
