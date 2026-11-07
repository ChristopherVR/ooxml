/**
 * group-drill-canvas.ts: the Angular canvas's side of selecting inside a group
 * the way PowerPoint does (shared `render/group-drill` makes every decision;
 * this only feeds it what `SlideCanvasComponent` has at hand and shapes the
 * answers for the template). No Angular dependency, so each rule is unit-testable
 * without mounting the component.
 *
 *   - what a press or a double-click selects is the shared
 *     `resolveGroupPressTarget` / `resolveGroupDoubleClickTarget`, which every
 *     binding calls;
 *   - Escape steps a selected member up to its parent group
 *     ({@link escapeSelectionTarget});
 *   - members are looked up in slide space next to the top-level elements
 *     ({@link withGroupMembers}), and the entered group gets a dashed frame
 *     ({@link enteredGroupBox}).
 *
 * Grouped children render with `pointer-events: none`, so the DOM hit-test
 * always answers with the top-level group; the member under the pointer is
 * found geometrically from the slide-space press point.
 */

import { parentSelection, slideSpaceElement, slideSpaceMembers } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Box } from './drag-resize';

/**
 * `elements` (the canvas's template + slide layers) plus every member of the
 * slide's enterable groups in slide space, so a member id resolves like a
 * top-level one for the selection chrome, drag, resize and the inline editor.
 * Top-level ids win; the same array comes back when there are no members.
 */
export function withGroupMembers(
	elements: readonly PptxElement[],
	slideElements: readonly PptxElement[],
): readonly PptxElement[] {
	const members = slideSpaceMembers(slideElements);
	if (members.size === 0) {
		return elements;
	}
	const known = new Set(elements.map((el) => el.id));
	const extra = [...members.values()].filter((member) => !known.has(member.id));
	return extra.length > 0 ? [...elements, ...extra] : elements;
}

/**
 * The slide-space box of the group a single selected member sits in (the
 * "entered" group, drawn with a dashed frame), or null when the selection is
 * not exactly one group member.
 */
export function enteredGroupBox(
	slideElements: readonly PptxElement[],
	selectedIds: readonly string[],
): (Box & { id: string }) | null {
	if (selectedIds.length !== 1) {
		return null;
	}
	const parentId = parentSelection(slideElements, selectedIds[0]);
	const group = parentId ? slideSpaceElement(slideElements, parentId) : null;
	return group
		? { id: group.id, x: group.x, y: group.y, width: group.width, height: group.height }
		: null;
}

/**
 * What Escape selects once no inline editor or transient chrome is left: the
 * parent group of a single selected member, or null (clear the selection).
 */
export function escapeSelectionTarget(
	slideElements: readonly PptxElement[] | undefined,
	selectedIds: readonly string[],
): string | null {
	if (!slideElements || selectedIds.length !== 1) {
		return null;
	}
	return parentSelection(slideElements, selectedIds[0]);
}
