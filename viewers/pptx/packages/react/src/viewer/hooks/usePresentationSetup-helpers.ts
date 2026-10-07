/**
 * Compatibility exports for presentation setup helpers that now live in the
 * framework-neutral shared package.
 */
export {
	applyRehearsalTimings,
	computeEntranceAnimationDelay,
	shouldLoopContinuously,
	sortEntranceAnimations,
} from 'ooxml-ui/pptx';
export type {
	EntranceAnimationEntry as AnimationEntry,
	PresentationLoopInput,
} from 'ooxml-ui/pptx';
