// Compatibility exports: document operations live in ooxml-core.
export {
	setTemplateElements,
	findTemplateElement,
	isElementIdInteractive,
	visibleTemplateElements,
	partitionTemplateElements,
	buildSaveSlides,
} from 'ooxml-core/pptx/editor/render/template-editing';
export type {
	TemplateElementMap,
	TemplateElementPartition,
} from 'ooxml-core/pptx/editor/render/template-editing';
