/** Whether `node` is `scope` or inside it, looking through shadow roots to their hosts. */
export function composedWithin(node: Node | undefined, scope: Node): boolean {
	let at: Node | null | undefined = node;
	while (at && at !== scope) at = at.parentNode ?? (at as Partial<ShadowRoot>).host;
	return at === scope;
}

/** The focused element, resolved through nested open shadow roots. */
export function deepActiveElement(root: Document | ShadowRoot): Element | null {
	let active = root.activeElement;
	while (active?.shadowRoot?.activeElement) {
		active = active.shadowRoot.activeElement;
	}
	return active;
}
