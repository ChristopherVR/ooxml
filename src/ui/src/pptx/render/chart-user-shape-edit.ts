// Compatibility exports: document operations live in ooxml-core.
export {
	listChartUserShapeDescriptors,
	createDefaultChartUserShape,
	pixelRectToRelAnchor,
	pixelRectToAbsAnchor,
	withChartUserShapeAdded,
	withChartUserShapeUpdated,
	withChartUserShapeRemoved,
	validateChartUserShapeAnchor,
} from 'ooxml-core/pptx/editor/render/chart-user-shape-edit';
export type { ChartUserShapeDescriptor } from 'ooxml-core/pptx/editor/render/chart-user-shape-edit';
