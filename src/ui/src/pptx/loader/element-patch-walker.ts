// Compatibility exports: document operations live in ooxml-core.
export {
	walkAndPatchElements,
	applyImagePathPatches,
} from 'ooxml-core/pptx/editor/loader/element-patch-walker';
export type { ElementPatcher } from 'ooxml-core/pptx/editor/loader/element-patch-walker';
