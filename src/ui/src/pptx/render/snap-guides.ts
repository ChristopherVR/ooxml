/** Compatibility entry: the implementation lives in ooxml-core/geometry (`snap-guides.ts`). */
export {
	SNAP_THRESHOLD,
	computeSnapToShape,
	computeSnap,
	snapToGridStep,
	snapValue,
	snapBox,
} from 'ooxml-core/geometry';
export type {
	SnapSibling,
	SnapGuideInput,
	SnapLine,
	SnapToShapeResult,
	SnapBox,
	SnapGuide,
	SnapResult,
} from 'ooxml-core/geometry';
