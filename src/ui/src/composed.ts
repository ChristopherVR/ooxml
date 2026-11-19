/** Whether `node` is `scope` or inside it, looking through shadow roots to their hosts. */
export function composedWithin(node: Node | undefined, scope: Node): boolean {
	let at: Node | null | undefined = node;
	while (at && at !== scope) at = at.parentNode ?? (at as Partial<ShadowRoot>).host;
	return at === scope;
}
