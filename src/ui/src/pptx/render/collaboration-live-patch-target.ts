// Compatibility exports: document operations live in ooxml-core.
export {
	GEOMETRY_KEYS,
	MAX_LIVE_TEXT_LENGTH,
	findElementYMap,
	applyLivePatch,
} from 'ooxml-core/pptx/editor/render/collaboration-live-patch-target';
export type {
	LiveGeometryPatch,
	LiveTextSource,
	PendingTextPatch,
	PendingPatch,
} from 'ooxml-core/pptx/editor/render/collaboration-live-patch-target';
