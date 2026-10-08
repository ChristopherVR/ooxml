/**
 * The `dgm:choose` algorithm resolution of `diagram/layout/
 * smartart-layout-interpreter-choose-algorithm.ts`, bound to the pptx raw
 * slots through {@link pptxOrderedXml}. See that module for the rules.
 */

import {
	chooseAlgorithm as chooseDiagramAlgorithm,
	chooseAlgorithmOfType as chooseDiagramAlgorithmOfType,
} from '../../../diagram/layout/smartart-layout-interpreter-choose-algorithm';
import type { WhenContext } from '../../../diagram/layout/smartart-layout-interpreter-when';
import type { PptxSmartArtLayoutAlgorithm, PptxSmartArtLayoutNode } from '../types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

/** The structural algorithm type a decidable `dgm:choose` on `node` selects. */
export function chooseAlgType(
	node: PptxSmartArtLayoutNode,
	nodeCount: number,
	context: WhenContext = {},
): string | undefined {
	return chooseDiagramAlgorithm(pptxOrderedXml, node, nodeCount, context)?.type;
}

/** The winning branch's full algorithm (type plus `dgm:param`s). */
export function chooseAlgorithm(
	node: PptxSmartArtLayoutNode,
	nodeCount: number,
	context: WhenContext = {},
): PptxSmartArtLayoutAlgorithm | undefined {
	return chooseDiagramAlgorithm(pptxOrderedXml, node, nodeCount, context);
}

/** The winning branch's algorithm when its type is one of `allowedTypes`. */
export function chooseAlgorithmOfType(
	node: PptxSmartArtLayoutNode,
	nodeCount: number,
	allowedTypes: ReadonlySet<string>,
	context: WhenContext = {},
): PptxSmartArtLayoutAlgorithm | undefined {
	return chooseDiagramAlgorithmOfType(pptxOrderedXml, node, nodeCount, allowedTypes, context);
}
