import { reviewParagraphFormatting, type ReviewDisplayMode } from 'ooxml-core/docx';
import { paragraphFromAttrs } from 'ooxml-core/docx/ui';
import type { Node } from 'prosemirror-model';

/** Shared model conversion keeps display and exported paragraph properties consistent. */
export function displayParagraph(node: Node, mode: ReviewDisplayMode) {
	return reviewParagraphFormatting(paragraphFromAttrs(node.attrs, String(node.attrs.id), []), mode);
}

/** Clear current direct declarations before applying a successfully projected prior paragraph. */
export const ORIGINAL_PARAGRAPH_RESET =
	'margin:0;margin-inline-start:0;margin-inline-end:0;text-indent:0;line-height:1.2;direction:ltr;border:none;background-color:transparent;float:none';
