import type { PptxElement } from 'ooxml-core/pptx';

import {
	drillSelectionForClick,
	drillSelectionForDoubleClick,
	findElementPath,
	isEnterableGroup,
	memberChainAtPoint,
} from './group-drill';

/**
 * What a press or a double-click on a top-level element selects, the way
 * PowerPoint selects inside a group. All five bindings ran the same sequence
 * of `group-drill` calls (enterable group, member chain under the point, the
 * selection's path, the drill decision); this is that sequence, once.
 *
 *   - the first press on a group selects the group; with it selected, a press
 *     selects the member under the pointer, and a press on another member of
 *     the entered group moves to that member;
 *   - a double-click goes straight to the innermost member under the pointer.
 *
 * Grouped children render with `pointer-events: none`, so the DOM hit-test
 * always answers with the top-level group; the member is found geometrically
 * from the slide-space press point. `memberChainAtPoint` already stops above a
 * `noSelect` member, so every id these return may be selected.
 */

/** A point in slide px (stage coordinates, zoom already divided out). */
export interface GroupPressPoint {
	x: number;
	y: number;
}

/**
 * The ids under `point` inside the top-level element `pressedId`, innermost
 * first and `pressedId` last; null when it isn't an enterable group (rotated,
 * flipped, empty, or not a group at all), so the press selects it as one.
 */
export function groupDrillChain(
	slideElements: readonly PptxElement[],
	pressedId: string,
	point: GroupPressPoint | null,
): string[] | null {
	const top = slideElements.find((el) => el.id === pressedId);
	if (!top || !point || !isEnterableGroup(top)) {
		return null;
	}
	return memberChainAtPoint(slideElements, pressedId, point);
}

/**
 * What a (non-additive) press on the top-level `pressedId` selects. Only a
 * single selection drills: with several elements selected, a press on one of
 * them keeps it whole so the selection can be dragged together.
 */
export function resolveGroupPressTarget(
	slideElements: readonly PptxElement[],
	pressedId: string,
	point: GroupPressPoint | null,
	selectedIds: readonly string[],
): string {
	const chain = groupDrillChain(slideElements, pressedId, point);
	if (!chain) {
		return pressedId;
	}
	const selectedId = selectedIds.length === 1 ? selectedIds[0] : undefined;
	const selectedPath = selectedId ? findElementPath(slideElements, selectedId) : null;
	return drillSelectionForClick(chain, selectedPath) ?? pressedId;
}

/** What a double-click on the top-level `pressedId` targets: the innermost member under the pointer. */
export function resolveGroupDoubleClickTarget(
	slideElements: readonly PptxElement[],
	pressedId: string,
	point: GroupPressPoint | null,
): string {
	const chain = groupDrillChain(slideElements, pressedId, point);
	return (chain ? drillSelectionForDoubleClick(chain) : null) ?? pressedId;
}
