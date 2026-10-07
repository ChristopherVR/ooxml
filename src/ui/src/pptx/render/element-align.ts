/** Compatibility entry: the implementation lives in ooxml-core/geometry (`align-distribute.ts`). */
export {
	usesSlideReference,
	resolveAlignReferenceBox,
	alignElements,
	distributeElements,
	computeAlign,
	computeDistribute,
} from 'ooxml-core/geometry';
export type {
	AlignEdge,
	DistributeAxis,
	AlignReference,
	AlignSlideSize,
	AlignOptions,
	ElementPosition,
	BoundingBoxElement,
	AlignReferenceBox,
	AlignMode,
	DistributeMode,
	AlignBox,
	PositionUpdate,
} from 'ooxml-core/geometry';
