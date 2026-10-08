/**
 * Folded-descendant text for the DiagramML interpreter bridge: which
 * descendants fold into a node's own box lives in `diagram/layout/
 * smartart-interpreter-fold-text.ts` (re-exported here); this module adds
 * the pptx text-run projection of a folded box.
 */

import type { PptxSmartArtNode, TextSegment, TextStyle } from '../types';
import { projectSmartArtNodeText } from './smartart-node-text-projection';

export {
	collectFoldedDescendants,
	foldedDescendantTexts,
	foldedItemText,
} from '../../../diagram/layout/smartart-interpreter-fold-text';

/**
 * Combine `node`'s own projected text with any {@link collectFoldedDescendants}
 * as additional paragraphs, matching the `\n`-joined text PowerPoint's own
 * cached drawing carries for the same case (`smartArtParagraphsText`).
 *
 * A folded descendant renders at `descendantFallbackStyle` (its OWN,
 * independently-shrunk size - see `smartart-layout-item-font-tier.ts`'s
 * module doc comment), which differs from `fallbackStyle` (the node's own
 * top-level size) whenever the arranger resolved one via
 * `RenderedNode.descendantFontSize`; falls back to `fallbackStyle` itself
 * when the arranger did not (every non-`lin`/`snake` caller of this bridge).
 */
export function projectFoldedNodeText(
	node: PptxSmartArtNode,
	folded: PptxSmartArtNode[],
	fallbackStyle: TextStyle,
	descendantFallbackStyle: TextStyle,
	bulletEnabled: boolean,
): { text: string; segments: TextSegment[] } {
	const segments = projectSmartArtNodeText(node, fallbackStyle, { bulletEnabled });
	const texts = [node.text];
	for (const descendant of folded) {
		segments.push({ text: '', style: descendantFallbackStyle, isParagraphBreak: true });
		segments.push(
			...projectSmartArtNodeText(descendant, descendantFallbackStyle, { bulletEnabled }),
		);
		texts.push(descendant.text);
	}
	return { text: texts.join('\n'), segments };
}
