// Compatibility exports: paragraph and run editing belongs to PowerPoint core.
// Named, not `export *`: a star re-export of an external module inside the bundle becomes a
// runtime namespace copy that keeps every `ooxml-ui/pptx` export reachable for a host bundler.
export {
	BODY_TEXT_STYLE_KEYS,
	isBodyTextStyleKey,
	overrideParagraphIndents,
	overrideParagraphProperties,
	paragraphOverridePatch,
	reconcileDecorationFlags,
	splitTextStyleUpdate,
	withExclusiveLineSpacing,
} from 'ooxml-core/pptx/editor/render/text-style-scope';
export type {
	ParagraphLevelPatch,
	TextStyleUpdateSplit,
} from 'ooxml-core/pptx/editor/render/text-style-scope';
