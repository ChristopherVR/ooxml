type Child = Node | string | null | undefined | false;

/** A small element factory, so no host data ever goes through innerHTML. */
export function h<K extends keyof HTMLElementTagNameMap>(
	tag: K,
	props: Partial<Record<string, string | ((event: Event) => void)>> = {},
	...children: Child[]
): HTMLElementTagNameMap[K] {
	const element = document.createElement(tag);
	for (const [key, value] of Object.entries(props)) {
		if (typeof value === 'function') element.addEventListener(key.replace(/^on/u, ''), value);
		else if (value !== undefined) element.setAttribute(key, value);
	}
	for (const child of children) if (child) element.append(child);
	return element;
}

export const time = (at: number): string =>
	new Date(at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
