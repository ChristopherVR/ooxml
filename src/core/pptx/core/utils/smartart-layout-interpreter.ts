/**
 * The DiagramML layout interpreter (`diagram/layout/
 * smartart-layout-interpreter.ts`) bound to the pptx raw slots through
 * {@link pptxOrderedXml}. See the diagram module for the pipeline.
 */

import {
	interpretSmartArtLayout as interpretDiagramLayout,
	type InterpretLayoutInput,
} from '../../../diagram/layout/smartart-layout-interpreter';
import type { SmartArtLayoutResult } from '../../../diagram/layout/smartart-layout-types';
import { pptxOrderedXml } from './smartart-ordered-xml-adapter';

export type { InterpretLayoutInput };

/** Interpret a SmartArt layout definition into positioned, styled geometry. */
export function interpretSmartArtLayout(
	input: InterpretLayoutInput,
): SmartArtLayoutResult | undefined {
	return interpretDiagramLayout(pptxOrderedXml, input);
}
