/**
 * The auxiliary arrangers (`diagram/layout/smartart-layout-interpreter-aux.ts`)
 * with `arrangeConn` bound to the pptx raw slots through
 * {@link pptxOrderedXml}.
 */

import type { ConstraintIndex } from '../../../diagram/layout/smartart-constraint-solver';
import { arrangeConn as arrangeDiagramConn } from '../../../diagram/layout/smartart-layout-interpreter-aux';
import type { ArrangementPlan } from '../../../diagram/layout/smartart-layout-interpreter-model';
import type {
	BoundingBox,
	SmartArtLayoutResult,
} from '../../../diagram/layout/smartart-layout-types';
import type { PptxSmartArtNode, SmartArtStyle } from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

export {
	arrangeSpacer,
	arrangeText,
} from '../../../diagram/layout/smartart-layout-interpreter-aux';

/** Connector-only arrangement (see the diagram module). */
export function arrangeConn(
	plan: ArrangementPlan,
	nodes: PptxSmartArtNode[],
	box: BoundingBox,
	palette: string[],
	style: SmartArtStyle,
	elementId: string,
	index?: ConstraintIndex,
): SmartArtLayoutResult | undefined {
	return arrangeDiagramConn(pptxOrderedXml, plan, nodes, box, palette, style, elementId, index);
}
