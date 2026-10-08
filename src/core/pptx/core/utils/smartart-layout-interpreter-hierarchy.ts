/**
 * The hierarchy arranger (`diagram/hierarchy/
 * smartart-layout-interpreter-hierarchy.ts`) bound to the pptx raw slots
 * through {@link pptxOrderedXml}. See the diagram module for the rules.
 */

import { arrangeHierarchy as arrangeDiagramHierarchy } from '../../../diagram/hierarchy/smartart-layout-interpreter-hierarchy';
import type { ConstraintIndex } from '../../../diagram/layout/smartart-constraint-solver';
import { EMPTY_CONSTRAINT_INDEX } from '../../../diagram/layout/smartart-constraint-solver';
import type {
	BoundingBox,
	SmartArtLayoutResult,
} from '../../../diagram/layout/smartart-layout-types';
import type {
	PptxSmartArtLayoutNode,
	PptxSmartArtNode,
	PptxSmartArtPresLayoutVars,
	SmartArtStyle,
} from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

/** Execute the hierarchy algorithm over the data-model node tree. */
export function arrangeHierarchy(
	nodes: PptxSmartArtNode[],
	box: BoundingBox,
	palette: string[],
	style: SmartArtStyle,
	elementId: string,
	presLayoutVars: PptxSmartArtPresLayoutVars | undefined,
	connectorLabels?: Map<string, string>,
	algorithmNode?: PptxSmartArtLayoutNode,
	index: ConstraintIndex = EMPTY_CONSTRAINT_INDEX,
	childOrder?: Map<string, number>,
	fontName?: string,
): SmartArtLayoutResult {
	return arrangeDiagramHierarchy(
		pptxOrderedXml,
		nodes,
		box,
		palette,
		style,
		elementId,
		presLayoutVars,
		connectorLabels,
		algorithmNode,
		index,
		childOrder,
		fontName,
	);
}
