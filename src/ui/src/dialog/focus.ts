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

/** A subtree that can hold no Tab stop. Slots and `display: contents` hosts must not count. */
function isHiddenBranch(element: HTMLElement): boolean {
	if (element.hidden || element.inert || element.getAttribute('aria-hidden') === 'true') {
		return true;
	}
	if (typeof element.checkVisibility === 'function') {
		return false;
	}
	const style = element.ownerDocument.defaultView?.getComputedStyle(element);
	return style?.display === 'none' || style?.visibility === 'hidden';
}

/** Whether a candidate is rendered, ancestors included. */
function isRendered(element: HTMLElement): boolean {
	return element.checkVisibility?.({ visibilityProperty: true }) ?? true;
}

/** What is rendered under `node`: slotted content for a slot, shadow content for a host. */
function renderedChildren(node: Element | ShadowRoot): Element[] {
	if (node instanceof Element && node.localName === 'slot') {
		const assigned = (node as HTMLSlotElement).assignedElements({ flatten: true });
		if (assigned.length > 0) return assigned;
	}
	const shadow = node instanceof Element ? node.shadowRoot : null;
	return Array.from((shadow ?? node).children);
}

/**
 * Tabbable elements in rendered order: the order Tab visits them, which is the flattened tree
 * (shadow content, with slotted children where their slot sits), not light-DOM order.
 */
function focusableElements(root: Element | ShadowRoot, found: HTMLElement[] = []): HTMLElement[] {
	for (const child of renderedChildren(root)) {
		if (child instanceof HTMLElement) {
			if (isHiddenBranch(child)) continue;
			if (child.matches(MODAL_FOCUSABLE_SELECTOR) && isRendered(child)) found.push(child);
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

/** Tabbable elements of `panel` in the order Tab visits them, reaching into open shadow roots. */
export function modalFocusableElements(panel: HTMLElement): HTMLElement[] {
	return focusableElements(panel);
}

/** The focused element of `doc`, resolved through nested open shadow roots. */
export function modalActiveElement(doc: Document): Element | null {
	return deepActiveElement(doc);
}

/** `panel.contains` across shadow boundaries and slots. */
export function containsComposed(panel: Element, node: Node | null): boolean {
	let current: Node | null = node;
	while (current) {
		if (current === panel) return true;
		current =
			(current as Element).assignedSlot ??
			current.parentNode ??
			(current as ShadowRoot).host ??
			null;
	}
	return false;
}
