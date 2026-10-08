/**
 * The `dgm:choose` structural-depth checks of `diagram/layout/
 * smartart-layout-interpreter-choose-depth.ts`, bound to the pptx raw slots
 * through {@link pptxOrderedXml}. See that module for the derivation.
 */

import {
	structuralChooseAlgDepth as diagramStructuralChooseAlgDepth,
	tunnelsPastOwnCompositeSlot as diagramTunnelsPastOwnCompositeSlot,
} from '../../../diagram/layout/smartart-layout-interpreter-choose-depth';
import type { StructuralChooseAlgResult } from '../../../diagram/layout/smartart-layout-interpreter-choose-depth';
import type { WhenContext } from '../../../diagram/layout/smartart-layout-interpreter-when';
import type { PptxSmartArtLayoutNode } from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

export type { StructuralChooseAlgResult } from '../../../diagram/layout/smartart-layout-interpreter-choose-depth';

/** Layout-node depth of the first structural `dgm:alg` in `node`'s winning choose branch. */
export function structuralChooseAlgDepth(
	node: PptxSmartArtLayoutNode,
	nodeCount: number,
	context: WhenContext,
): StructuralChooseAlgResult | undefined {
	return diagramStructuralChooseAlgDepth(pptxOrderedXml, node, nodeCount, context);
}

/** Whether `node` is its own composite and its choose only tunnels into one of its slots. */
export function tunnelsPastOwnCompositeSlot(
	node: PptxSmartArtLayoutNode,
	nodeCount: number,
	context: WhenContext,
	itemTemplates: ReadonlySet<PptxSmartArtLayoutNode>,
): boolean {
	return diagramTunnelsPastOwnCompositeSlot(
		pptxOrderedXml,
		node,
		nodeCount,
		context,
		itemTemplates,
	);
}
