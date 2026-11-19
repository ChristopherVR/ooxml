/**
 * Keeping keyboard focus where it was while an editor redraws its document (an assistant edit or
 * its undo reloads the file). The element that had focus is found again by its shape id, or by its
 * role and name, so a screen reader user stays on the same shape or control.
 */

/** The focused element, followed through open shadow roots. */
export function deepActive() {
	let active = document.activeElement;
	while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
	return active;
}

/** Whether `node` holds `target`, across shadow roots. */
export function holds(node, target) {
	for (let at = target; at; at = at.parentNode ?? at.host) if (at === node) return true;
	return false;
}

function* deep(root) {
	for (const element of root.querySelectorAll('*')) {
		yield element;
		if (element.shadowRoot) yield* deep(element.shadowRoot);
	}
}

/** Where focus is inside `node`, or null when it is elsewhere. */
export function focusPlace(node) {
	const element = deepActive();
	if (!element || !holds(node, element)) return null;
	return {
		element,
		tag: element.localName,
		shape: element.dataset?.shapeId,
		role: element.getAttribute('role'),
		label: element.getAttribute('aria-label'),
	};
}

/** Puts focus back on the element of `place`, or its redrawn twin. True when focus is in `node`. */
export function returnFocus(node, place) {
	if (holds(node, deepActive())) return true;
	const same = (element) =>
		element.localName === place.tag &&
		(place.shape
			? element.dataset?.shapeId === place.shape
			: element.getAttribute('role') === place.role &&
				element.getAttribute('aria-label') === place.label);
	const target = place.element.isConnected && holds(node, place.element) ? place.element : null;
	(target ?? [...deep(node)].find(same))?.focus({ preventScroll: true });
	return holds(node, deepActive());
}
