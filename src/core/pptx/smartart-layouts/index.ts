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
} from '../core/utils/smartart-builtin-layouts';
export type { BuiltinSmartArtLayoutEntry } from '../core/utils/smartart-builtin-layouts';
