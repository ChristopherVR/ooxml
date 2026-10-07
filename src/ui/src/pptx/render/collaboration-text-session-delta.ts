// Compatibility exports: document operations live in ooxml-core.
export {
	textSessionReplacementSpan,
	textSessionRetainedIndices,
	textSessionPlan,
	textSessionAttributesEqual,
	textSessionDeltaSupported,
	textSessionUnits,
	textSessionAttributePatch,
	mergeTextSessionObjectAttribute,
} from 'ooxml-core/pptx/editor/render/collaboration-text-session-delta';
export type {
	TextSessionUnit,
	LocalTextReplacement,
	TextSessionCorrespondence,
	LocalTextEdit,
} from 'ooxml-core/pptx/editor/render/collaboration-text-session-delta';
