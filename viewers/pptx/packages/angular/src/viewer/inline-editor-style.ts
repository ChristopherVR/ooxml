import type { PptxElement } from 'pptx-viewer-core';

import { getTextBlockStyle } from './element-style';
import type { StyleMap } from './element-style';

/**
 * The inline text editor's style, for the plain editor and the list editor:
 * the block style view mode paints the text with (alignment, anchor, insets,
 * colour and the substituted, autofit-scaled font), placed on the element's
 * box in stage coordinates.
 *
 * The editor used to be an unstyled `<textarea>` (and the list editor an
 * unstyled surface) with opaque white paper and plain black text, so the text
 * changed font, size, colour and position the moment it was double-clicked.
 * The runs carry their own style (`seedPlainInlineText` / the list DOM).
 */
export function inlineEditorStyle(element: PptxElement, rich = false): StyleMap {
	return {
		...getTextBlockStyle(element),
		left: `${element.x}px`,
		top: `${element.y}px`,
		width: `${element.width}px`,
		height: `${element.height}px`,
		// The list and collaborative editors put their paragraphs straight in the
		// surface, where flex items would each become a row.
		...(rich
			? { display: 'block', 'text-decoration': 'none', 'text-decoration-line': 'none' }
			: {}),
	};
}
