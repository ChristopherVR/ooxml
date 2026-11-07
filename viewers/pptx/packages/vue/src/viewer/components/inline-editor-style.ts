import type { PptxElement } from 'pptx-viewer-core';
import type { CSSProperties } from 'vue';

import { getTextBlockStyle } from '../composables/element-style';

/**
 * The inline editor's style: the block style view mode paints the text with
 * (alignment, anchor, insets, colour and the substituted, autofit-scaled font),
 * placed in the slide's own coordinate space.
 *
 * This used to copy a few raw `textStyle` fields instead. The raw `fontFamily`
 * has no substitution chain, so a deck font that was not painted (a webfont
 * still downloading or blocked, offline, or the download option off) fell back
 * to the browser's default serif here while view mode used the chain's
 * stand-in, and the raw size skipped the autofit scale: the text changed font
 * and size on double-click. The runs carry their own style (`seedPlainInlineText`).
 */
export function inlineEditorStyle(el: PptxElement, rich: boolean): CSSProperties {
	return {
		...getTextBlockStyle(el),
		position: 'absolute',
		left: `${el.x}px`,
		top: `${el.y}px`,
		width: `${el.width}px`,
		height: `${el.height}px`,
		// The list and collaborative editors put their paragraphs straight in the
		// surface, where flex items would each become a row.
		...(rich ? { display: 'block', textDecoration: 'none', textDecorationLine: 'none' } : {}),
		boxSizing: 'border-box',
		margin: 0,
		outline: '2px solid var(--pptx-vue-selection-color, #3b82f6)',
		// Transparent: the element's shape fill remains rendered underneath. Only
		// its static text is suppressed, so non-white authored fills stay visible.
		background: 'transparent',
		overflow: 'hidden',
		whiteSpace: 'pre-wrap',
		cursor: 'text',
		zIndex: 60,
	};
}
