/**
 * Keyboard focus in a slide rail. PowerPoint keeps focus on the selected
 * thumbnail as the user pages through slides; a virtual window removes the
 * row that held focus once it scrolls out, which dropped focus to the page
 * and stopped keyboard paging in the rail.
 */

/** True when keyboard focus is inside the rail list. */
export function railHasFocus(list: Element | null | undefined): boolean {
	const active = list?.ownerDocument.activeElement;
	return Boolean(list && active && list.contains(active));
}

/** Move focus to the current slide's row, without scrolling the rail. */
export function focusCurrentRow(list: Element | null | undefined): void {
	const row = list?.querySelector<HTMLElement>(
		'[data-pptx-chrome="slide-row"][aria-current="true"]',
	);
	if (row && row.ownerDocument.activeElement !== row) {
		row.focus({ preventScroll: true });
	}
}
