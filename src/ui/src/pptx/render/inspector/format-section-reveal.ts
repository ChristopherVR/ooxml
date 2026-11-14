/**
 * Bring a Format pane section into view after a launcher asked for it, and
 * map the older context-menu anchors onto section ids.
 *
 * @module render/inspector/format-section-reveal
 */
import type { InspectorSectionAnchor } from '../context-menu-inspector-anchor';
import { FORMAT_SECTION_ATTRIBUTE } from './format-sections';
import type { FormatSectionId } from './format-sections';

/** The section a context-menu anchor opens (Format Shape: Fill, Size and Position, Alt Text). */
export function formatSectionForAnchor(anchor: InspectorSectionAnchor): FormatSectionId {
	switch (anchor) {
		case 'transform':
			return 'sizePosition';
		case 'alt-text':
			return 'altText';
		default:
			return 'fill';
	}
}

/** The section element of `section` inside `root`, when rendered. */
export function findFormatSection(root: ParentNode, section: FormatSectionId): HTMLElement | null {
	return root.querySelector<HTMLElement>(`[${FORMAT_SECTION_ATTRIBUTE}="${section}"]`);
}

/**
 * Scroll `section` into view and, with `focus`, move focus to its heading. A
 * section that is not rendered (hidden by the host, or not mounted yet) is
 * skipped; returns whether it was found.
 */
export function revealFormatSection(
	root: ParentNode,
	section: FormatSectionId,
	options: { focus?: boolean } = {},
): boolean {
	const node = findFormatSection(root, section);
	if (!node) {
		return false;
	}
	node.scrollIntoView?.({ block: 'start' });
	if (options.focus) {
		node.focus({ preventScroll: true });
	}
	return true;
}
