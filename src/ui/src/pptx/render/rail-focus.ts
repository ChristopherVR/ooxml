/**
 * Keyboard focus in a slide rail. PowerPoint keeps focus on the selected
 * thumbnail as the user pages through slides; a virtual window removes the
 * row that held focus once it scrolls out, which dropped focus to the page
 * and stopped keyboard paging in the rail.
 */

const RAIL = '[data-pptx-chrome="slide-list"]';
/** The rail of each document that keyboard focus last entered. */
const lastRail = new WeakMap<Document, Element>();
const tracked = new WeakSet<Document>();

/**
 * Notes which rail focus enters. A render that removes the focused row (an
 * undone New Slide, a rebuilt rail) drops focus to the page without focus
 * landing anywhere else, so the rail keeps it until focus or a press goes
 * elsewhere.
 */
function track(doc: Document): void {
	if (tracked.has(doc)) {
		return;
	}
	tracked.add(doc);
	const note = (event: Event): void => {
		const target = event.composedPath()[0] as Element | undefined;
		const rail = target?.closest?.(RAIL);
		if (rail) {
			lastRail.set(doc, rail);
		} else {
			lastRail.delete(doc);
		}
	};
	doc.addEventListener('focusin', note, true);
	doc.addEventListener('pointerdown', note, true);
}

/**
 * True when keyboard focus is inside the rail list, or was until a render
 * removed the row that held it.
 */
export function railHasFocus(list: Element | null | undefined): boolean {
	if (!list) {
		return false;
	}
	const doc = list.ownerDocument;
	track(doc);
	const active = doc.activeElement;
	if (active && list.contains(active)) {
		return true;
	}
	return (!active || active === doc.body) && lastRail.get(doc) === list;
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
