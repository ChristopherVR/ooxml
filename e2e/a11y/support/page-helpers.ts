/**
 * Helpers installed in every page before the app starts (`addInitScript`), so the specs can ask
 * where keyboard focus is through shadow roots and slots. Runs in the browser: it must not
 * reference anything outside itself.
 */
export function installA11yHelpers(): void {
	const ids = new WeakMap<Element, number>();
	let next = 0;
	const up = (node: Node): Node | null =>
		(node as Element).assignedSlot ?? node.parentNode ?? (node as ShadowRoot).host ?? null;
	const deep = (): Element | null => {
		let active = document.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		return active && active !== document.body && active !== document.documentElement
			? active
			: null;
	};
	const matchesUp = (element: Element, selector: string): boolean => {
		for (let node: Node | null = element; node; node = up(node))
			if (node.nodeType === 1 && (node as Element).matches(selector)) return true;
		return false;
	};
	const name = (element: Element): string =>
		(
			element.getAttribute('aria-label') ||
			element.getAttribute('title') ||
			element.getAttribute('placeholder') ||
			element.textContent ||
			''
		)
			.trim()
			.replace(/\s+/gu, ' ')
			.slice(0, 40);
	const idOf = (element: Element): number => {
		if (!ids.has(element)) ids.set(element, ++next);
		return ids.get(element) ?? 0;
	};
	/** What a stop is to the user: a Visio shape stays itself when an edit draws it again. */
	const keyOf = (element: Element): string => {
		const { shapeId, pageId } = (element as SVGElement).dataset ?? {};
		return shapeId ? `shape ${pageId ?? ''}/${shapeId}` : `element ${idOf(element)}`;
	};
	const helpers = {
		deep,
		/**
		 * One focus stop: a stable element id, what it is to the user (`key`), a readable label and
		 * the first region it is in.
		 */
		stop(regions: [string, string][]) {
			const element = deep();
			if (!element) return { id: 0, key: 'body', label: 'body', region: null as string | null };
			const role = element.getAttribute('role');
			return {
				id: idOf(element),
				key: keyOf(element),
				label: `${element.localName}${role ? `[${role}]` : ''} "${name(element)}"`,
				region: regions.find(([, selector]) => matchesUp(element, selector))?.[0] ?? null,
			};
		},
		/** Whether focus is inside an element matching `selector`, through shadow roots and slots. */
		focusWithin: (selector: string) => {
			const element = deep();
			return element ? matchesUp(element, selector) : false;
		},
		/** Ids of the laid-out, visible elements matching `selector`, shadow roots included. */
		visible: (selector: string) =>
			helpers
				.queryAll(selector)
				.filter((element) => {
					const box = element.getBoundingClientRect();
					return (
						box.width > 0 && box.height > 0 && getComputedStyle(element).visibility !== 'hidden'
					);
				})
				.map(idOf),
		/** Every element matching `selector` in the document and its open shadow roots. */
		queryAll(selector: string, root: ParentNode = document): Element[] {
			const found = [...root.querySelectorAll(selector)];
			for (const element of root.querySelectorAll('*'))
				if (element.shadowRoot) found.push(...helpers.queryAll(selector, element.shadowRoot));
			return found;
		},
	};
	Object.defineProperty(globalThis, '__a11y', { value: helpers, configurable: true });
}

export interface FocusStop {
	id: number;
	key: string;
	label: string;
	region: string | null;
}

export interface A11yHelpers {
	deep(): Element | null;
	stop(regions: [string, string][]): FocusStop;
	focusWithin(selector: string): boolean;
	visible(selector: string): number[];
	queryAll(selector: string, root?: ParentNode): Element[];
}

declare global {
	var __a11y: A11yHelpers;
}
