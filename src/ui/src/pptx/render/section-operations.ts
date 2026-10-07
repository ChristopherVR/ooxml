// Compatibility exports: document operations live in ooxml-core.
export {
	generateSectionId,
	resolveSlideId,
	groupSlidesBySection,
	addSection,
	renameSection,
	deleteSection,
	moveSectionUp,
	moveSectionDown,
	moveSlidesToSection,
} from 'ooxml-core/pptx/editor/render/section-operations';
export type {
	SectionMutationResult,
	SectionGroupDescriptor,
	SectionSlideGroup,
} from 'ooxml-core/pptx/editor/render/section-operations';
