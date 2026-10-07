// Compatibility exports: document operations live in ooxml-core.
export {
	DEFAULT_BULLET_CHAR,
	DEFAULT_AUTONUM_TYPE,
	bulletInfoForKind,
	isBulletMarkerSegment,
	paragraphBulletKind,
	withoutListType,
	toggleParagraphBullet,
	splitBulletParagraphs,
	resolveBulletSegments,
	paragraphsBulletKind,
	elementBulletKind,
	setElementBullets,
	toggleElementBullets,
} from 'ooxml-core/pptx/editor/render/bullet-toggle';
export type {
	ParagraphBulletKind,
	ElementBulletKind,
	BulletParagraphRange,
	ElementBulletPatch,
	BulletParagraph,
} from 'ooxml-core/pptx/editor/render/bullet-toggle';
