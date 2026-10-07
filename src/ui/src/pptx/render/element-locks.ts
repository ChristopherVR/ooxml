// Compatibility exports: document operations live in ooxml-core.
export {
	getElementLocks,
	resolveElementInteractivity,
	canInteractWithElement,
	canDrillDown,
	isElementInteractionLocked,
	isElementLocked,
	elementLockTogglePatch,
	filterInteractableIds,
} from 'ooxml-core/pptx/editor/render/element-locks';
export type {
	ElementInteraction,
	ElementInteractivity,
} from 'ooxml-core/pptx/editor/render/element-locks';
