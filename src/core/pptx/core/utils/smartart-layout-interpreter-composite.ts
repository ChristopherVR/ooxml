/**
 * The composite arranger (`diagram/layout/
 * smartart-layout-interpreter-composite.ts`) bound to the pptx raw slots
 * through {@link pptxOrderedXml}.
 */

import type { ConstraintIndex } from '../../../diagram/layout/smartart-constraint-solver';
import { arrangeComposite as arrangeDiagramComposite } from '../../../diagram/layout/smartart-layout-interpreter-composite';
import type { ArrangementPlan } from '../../../diagram/layout/smartart-layout-interpreter-model';
import type {
	BoundingBox,
	SmartArtLayoutResult,
} from '../../../diagram/layout/smartart-layout-types';
import type { PptxSmartArtNode, PptxSmartArtPresLayoutVars, SmartArtStyle } from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

/** Place `nodes` into a composite layout's declared slots (see the diagram module). */
export function arrangeComposite(
	plan: ArrangementPlan,
	nodes: PptxSmartArtNode[],
	box: BoundingBox,
	palette: string[],
	style: SmartArtStyle,
	elementId: string,
	index?: ConstraintIndex,
	childrenOf?: Map<string, PptxSmartArtNode[]>,
	flat?: PptxSmartArtNode[],
	fontName?: string,
	presLayoutVars?: PptxSmartArtPresLayoutVars,
): SmartArtLayoutResult | undefined {
	return arrangeDiagramComposite(
		pptxOrderedXml,
		plan,
		nodes,
		box,
		palette,
		style,
		elementId,
		index,
		childrenOf,
		flat,
		fontName,
		presLayoutVars,
	);
}
