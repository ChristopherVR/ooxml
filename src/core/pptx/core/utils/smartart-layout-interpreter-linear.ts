/**
 * The linear arranger (`diagram/layout/smartart-layout-interpreter-linear.ts`)
 * bound to the pptx raw slots through {@link pptxOrderedXml}.
 */

import type { ConstraintIndex } from '../../../diagram/layout/smartart-constraint-solver';
import { arrangeLinear as arrangeDiagramLinear } from '../../../diagram/layout/smartart-layout-interpreter-linear';
import type {
	ArrangementPlan,
	FlowDirection,
} from '../../../diagram/layout/smartart-layout-interpreter-model';
import type {
	BoundingBox,
	SmartArtLayoutResult,
} from '../../../diagram/layout/smartart-layout-types';
import type { PptxSmartArtNode, PptxSmartArtPresLayoutVars, SmartArtStyle } from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

export { arrangeSnake } from '../../../diagram/layout/smartart-layout-interpreter-snake';

/** Lay `nodes` out along one axis (see the diagram module). */
export function arrangeLinear(
	plan: ArrangementPlan,
	flow: FlowDirection,
	nodes: PptxSmartArtNode[],
	box: BoundingBox,
	palette: string[],
	style: SmartArtStyle,
	elementId: string,
	index?: ConstraintIndex,
	childrenOf?: Map<string, PptxSmartArtNode[]>,
	fontName?: string,
	presLayoutVars?: PptxSmartArtPresLayoutVars,
): SmartArtLayoutResult {
	return arrangeDiagramLinear(
		pptxOrderedXml,
		plan,
		flow,
		nodes,
		box,
		palette,
		style,
		elementId,
		index,
		childrenOf,
		fontName,
		presLayoutVars,
	);
}
