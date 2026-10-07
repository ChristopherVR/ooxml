/** Shared DOM focus utilities for dialogs and other modal surfaces. */

export const MODAL_FOCUSABLE_SELECTOR = [
	'a[href]',
	'button:not([disabled])',
	'input:not([disabled]):not([type="hidden"])',
	'select:not([disabled])',
	'textarea:not([disabled])',
	'[contenteditable="true"]',
	'[tabindex]:not([tabindex="-1"])',
].join(',');

function isAvailable(element: HTMLElement): boolean {
	if (element.hidden || element.getAttribute('aria-hidden') === 'true') {
		return false;
	}
	const style = element.ownerDocument.defaultView?.getComputedStyle(element);
	return style?.display !== 'none' && style?.visibility !== 'hidden';
}

/**
 * Tabbable elements in tree order, including those inside open shadow roots
 * (the shared `pptx-ui-*` controls render their buttons there).
 */
function focusableElements(root: ParentNode, found: HTMLElement[] = []): HTMLElement[] {
	for (const child of Array.from(root.children)) {
		if (child instanceof HTMLElement) {
			if (!isAvailable(child)) continue;
			if (child.matches(MODAL_FOCUSABLE_SELECTOR) && isAvailable(child)) {
				found.push(child);
			}
			if (child.shadowRoot && isAvailable(child)) {
				focusableElements(child.shadowRoot, found);
			}
		}
		focusableElements(child, found);
	}
	return found;
}

/** The focused element, resolved through nested open shadow roots. */
function deepActiveElement(doc: Document): Element | null {
	let active = doc.activeElement;
	while (active?.shadowRoot?.activeElement) {
		active = active.shadowRoot.activeElement;
	}
	return active;
}

/** Tabbable elements of `panel` in tree order, reaching into open shadow roots. */
export function modalFocusableElements(panel: HTMLElement): HTMLElement[] {
	return focusableElements(panel);
}

/** The focused element of `doc`, resolved through nested open shadow roots. */
export function modalActiveElement(doc: Document): Element | null {
	return deepActiveElement(doc);
}
