/**
 * {@link LinkInput} → the model fields the writer saves: a shape-level
 * `PptxAction` (`p:cNvPr/a:hlinkClick`) or the run-level hyperlink fields of a
 * `TextStyle` (`a:rPr/a:hlinkClick`). A slide jump carries only its slide
 * index; the save adds the `…/relationships/slide` relationship for it, and
 * reports `SAVE_SLIDE_JUMP_TARGET_MISSING` when the index names no slide.
 *
 * @module sdk/link-input
 */

import type { PptxAction } from '../../types/actions';
import type { TextStyle } from '../../types/text';
import type { LinkInput } from './types';

const SLIDE_JUMP_ACTION = 'ppaction://hlinksldjump';

/** The click action of an element, or `undefined` without a link. */
export function mapLinkToAction(link: LinkInput | undefined): PptxAction | undefined {
	if (!link) {
		return undefined;
	}
	const tooltip = link.tooltip !== undefined ? { tooltip: link.tooltip } : {};
	if ('url' in link) {
		return { url: assertUrl(link.url), ...tooltip };
	}
	assertSlideIndex(link.slideIndex);
	return { action: SLIDE_JUMP_ACTION, targetSlideIndex: link.slideIndex, ...tooltip };
}

/** The run-level hyperlink fields of a text style; empty without a link. */
export function mapLinkToRunStyle(link: LinkInput | undefined): Partial<TextStyle> {
	if (!link) {
		return {};
	}
	const tooltip = link.tooltip !== undefined ? { hyperlinkTooltip: link.tooltip } : {};
	if ('url' in link) {
		return { hyperlink: assertUrl(link.url), ...tooltip };
	}
	assertSlideIndex(link.slideIndex);
	return {
		hyperlink: SLIDE_JUMP_ACTION,
		hyperlinkAction: SLIDE_JUMP_ACTION,
		hyperlinkTargetSlideIndex: link.slideIndex,
		...tooltip,
	};
}

function assertUrl(url: string): string {
	if (typeof url !== 'string' || url.trim() === '') {
		throw new RangeError('LinkInput.url must be a non-empty address');
	}
	return url;
}

function assertSlideIndex(index: number): void {
	if (!Number.isInteger(index) || index < 0) {
		throw new RangeError(`LinkInput.slideIndex must be a 0-based slide position, got ${index}`);
	}
}
