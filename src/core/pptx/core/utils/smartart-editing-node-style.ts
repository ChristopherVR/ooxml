/**
 * SmartArt per-node visual override, bound to `PptxSmartArtData`. The
 * operation lives in `diagram/edit/node-style`.
 *
 * @module smartart-editing-node-style
 */

import { setDiagramNodeStyle } from '../../../diagram/edit/index';
import type { PptxSmartArtData, PptxSmartArtNodeStyle } from '../types';

/** Set (merge) a per-node visual override on a SmartArt node by id. */
export function setSmartArtNodeStyle(
	data: PptxSmartArtData,
	nodeId: string,
	style: Partial<PptxSmartArtNodeStyle>,
): PptxSmartArtData {
	return setDiagramNodeStyle(data, nodeId, style);
}
