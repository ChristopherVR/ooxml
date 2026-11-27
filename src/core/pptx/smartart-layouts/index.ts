/**
 * `ooxml-core/pptx/smartart-layouts`: PowerPoint's built-in SmartArt layout definitions.
 *
 * A separate entry because the library is about 350 KB gzipped: the main `pptx` entry never
 * imports it, so only a viewer that offers a named-layout gallery pays for it.
 */
export {
	applyBuiltinSmartArtLayout,
	defaultBuiltinSmartArtLayoutId,
	findBuiltinSmartArtLayout,
	listBuiltinSmartArtLayouts,
	loadBuiltinSmartArtLayoutXml,
	parseBuiltinLayoutDefinition,
	findBuiltinSmartArtStyleDefinition,
	loadBuiltinSmartArtStyleDefinitionXml,
} from '../core/utils/smartart-builtin-layouts';
export type { BuiltinSmartArtLayoutEntry } from '../core/utils/smartart-builtin-layouts';
// Picture Format > Picture Layout: pictures into a built-in picture SmartArt layout.
export {
	listPictureLayouts,
	pictureLayoutSmartArt,
	PENDING_PICTURE_LAYOUT_PART,
	type PictureLayoutOptions,
	type PictureLayoutRefusal,
} from '../core/utils/smartart-builtin-layouts/picture-layout';
