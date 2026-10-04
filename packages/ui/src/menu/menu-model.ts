/**
 * Menu data and the pure helpers both menu modes share. Moved from pptx-viewer
 * (`web-components/context-menu-model.ts`, `render/flyout-position.ts`).
 */

/** One row of a controlled menu. Hosts decide `id`, wording, gating and what activation does. */
export interface OfficeMenuItem {
	/** Stable command id, echoed by the request event. */
	id: string;
	/** Translated visible label. */
	label: string;
	/** Draw a rule before this row. */
	separatorBefore?: boolean;
	/** Group heading shown before this row; the row and the ones after it join that group. */
	heading?: string;
	/** Destructive command (Delete): tinted. */
	danger?: boolean;
	/** Offered but unavailable: announced and drawn dim, skipped by the arrow keys. */
	disabled?: boolean;
	/** A toggle (Grid, Ruler) in this state; omitted for one-shot commands. */
	checked?: boolean;
	/** A registered icon drawn before the label. */
	icon?: string;
}

/** Everything a controlled menu needs to draw. */
export interface OfficeMenuState {
	/** Pointer position, in viewport pixels; the menu clamps itself into the window. */
	x: number;
	y: number;
	/** Accessible name. */
	label: string;
	items: readonly OfficeMenuItem[];
	/** Test-hook attributes copied onto the host (names matching the element's `markerPattern`). */
	markers?: readonly string[];
	/** Stacking order; the `--office-z-popover` token applies when omitted. */
	zIndex?: number;
	/** Move focus onto the first enabled command on open. Default true. */
	autoFocus?: boolean;
}

export type OfficeMenuCloseReason = 'escape' | 'outside' | 'tab';

export const EMPTY_MENU_STATE: OfficeMenuState = { x: 0, y: 0, label: '', items: [] };

/** Index of the next enabled row from `from` in `step` direction, wrapping; -1 if none. */
export function nextEnabledIndex(
	items: readonly { disabled?: boolean }[],
	from: number,
	step: 1 | -1,
): number {
	const count = items.length;
	for (let offset = 1; offset <= count; offset += 1) {
		const index = (((from + step * offset) % count) + count) % count;
		if (!items[index]?.disabled) return index;
	}
	return -1;
}

/**
 * Type-ahead target: the next enabled row after `from` whose label starts with `query`.
 * Repeating one character cycles through the rows that start with it.
 */
export function typeAheadIndex(
	items: readonly { label: string; disabled?: boolean }[],
	from: number,
	query: string,
): number {
	const needle = query.toLowerCase();
	const cycling = [...needle].every((char) => char === needle[0]);
	const prefix = cycling ? needle[0] : needle;
	for (let offset = cycling ? 1 : 0; offset <= items.length; offset += 1) {
		const index = (from + offset + items.length) % items.length;
		const item = items[index];
		if (
			item &&
			!item.disabled &&
			item.label
				.trim()
				.toLowerCase()
				.startsWith(prefix ?? '')
		)
			return index;
	}
	return -1;
}

/** Where a flyout wants to sit, how big it is, and what it must fit inside. */
export interface FlyoutPositionInput {
	/** Anchor point in viewport coordinates (a pointer event's clientX/Y). */
	x: number;
	y: number;
	/** Measured size; zero means "not measured yet". */
	width: number;
	height: number;
	/** The box it must stay inside, normally the window's inner size. */
	viewportWidth: number;
	viewportHeight: number;
	/** Gap kept from every edge, in px. Default 8. */
	margin?: number;
}

/**
 * The top-left corner keeping a pointer-anchored flyout fully on screen. Flips inwards at the
 * right and bottom edges, as desktop menus do, and never returns a negative coordinate, so a
 * flyout larger than the viewport stays reachable from the top.
 */
export function clampFlyoutPosition(input: FlyoutPositionInput): { left: number; top: number } {
	const margin = input.margin ?? 8;
	const maxLeft = input.viewportWidth - input.width - margin;
	const maxTop = input.viewportHeight - input.height - margin;
	return {
		left: Math.max(margin, input.width > 0 ? Math.min(input.x, maxLeft) : input.x),
		top: Math.max(margin, input.height > 0 ? Math.min(input.y, maxTop) : input.y),
	};
}
