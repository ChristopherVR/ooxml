/**
 * The layout interpreter's arrangement discovery (`diagram/layout/
 * smartart-layout-interpreter-model.ts`) bound to the pptx raw slots through
 * {@link pptxOrderedXml}, plus the helpers that module re-exports, so the
 * pptx interpreters keep one import site. See the diagram module for the
 * precedence rules.
 */

import { discoverArrangement as discoverDiagramArrangement } from '../../../diagram/layout/smartart-layout-interpreter-model';
import type { ArrangementPlan as DiagramArrangementPlan } from '../../../diagram/layout/smartart-layout-interpreter-model';
import type {
	PptxSmartArtLayoutDefinition,
	PptxSmartArtNode,
	PptxSmartArtPresLayoutVars,
	XmlObject,
} from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

export {
	algorithmParam,
	clampByRules,
	detectPositionFamily,
	findConstraint,
	hasPositionGuard,
	isMeaningfulAux,
	itemNode,
	numericParam,
	PRIMARY_ALG,
	ratioConstraint,
	resolveFlowDirection,
	STRUCTURAL_ARRANGEMENT_KINDS,
	type ArrangementKind,
	type FlowDirection,
} from '../../../diagram/layout/smartart-layout-interpreter-model';

/** The arranger a pptx layout definition resolves to (its node keeps the pptx raw slots). */
export type ArrangementPlan = DiagramArrangementPlan<XmlObject>;

/** Which arrangement algorithm drives the diagram (see the diagram module). */
export function discoverArrangement(
	definition: PptxSmartArtLayoutDefinition,
	nodeCount?: number,
	presLayoutVars?: PptxSmartArtPresLayoutVars,
	flatNodes?: PptxSmartArtNode[],
): ArrangementPlan | undefined {
	return discoverDiagramArrangement(
		pptxOrderedXml,
		definition,
		nodeCount,
		presLayoutVars,
		flatNodes,
	);
}
