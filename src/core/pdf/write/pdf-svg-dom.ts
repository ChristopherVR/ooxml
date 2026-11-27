/**
 * The DOM subset the SVG to PDF converter walks, kept to what both a browser DOM and the DOM-free
 * `@xmldom/xmldom` tree (`ooxml-core/xml`) answer: `localName`, attributes, `childNodes`,
 * `parentElement` and `textContent`. Browser-only conveniences (`children`, `querySelectorAll`,
 * `id`) are not used, so a drawing parsed without a browser (Node, a worker) converts the same way.
 */
const ELEMENT_NODE = 1;

/** The child elements of `node`, in document order (`Element.children`). */
export function childElements(node: Element): Element[] {
	const out: Element[] = [];
	for (const child of Array.from(node.childNodes))
		if (child.nodeType === ELEMENT_NODE) out.push(child as Element);
	return out;
}

/** Every element below `root`, in document order, without `root` itself. */
export function descendantElements(root: Element): Element[] {
	const out: Element[] = [];
	const stack = childElements(root).reverse();
	while (stack.length) {
		const node = stack.pop()!;
		out.push(node);
		const children = childElements(node);
		for (let index = children.length - 1; index >= 0; index--) stack.push(children[index]!);
	}
	return out;
}

/** An element's `id` attribute (`Element.id`): empty when it has none. */
export const elementId = (element: Element): string => element.getAttribute('id') ?? '';

/**
 * Parses SVG markup into its root `<svg>` element: with the host's `DOMParser` where there is
 * one (a browser), otherwise with the DOM-free XML parser. `undefined` when the markup is not
 * well-formed or its root is not `<svg>`. The XML parser is loaded only when it is needed, so a
 * browser bundle never downloads it for this.
 */
export async function parseSvgMarkup(markup: string): Promise<Element | undefined> {
	let root: Element | null | undefined;
	if (typeof DOMParser !== 'undefined') {
		const parsed = new DOMParser().parseFromString(markup, 'image/svg+xml');
		root = parsed.documentElement;
	} else {
		try {
			const { parseXml } = await import('ooxml-core/xml');
			root = parseXml(markup, { label: 'SVG' }).documentElement;
		} catch {
			return undefined;
		}
	}
	return root?.localName === 'svg' ? root : undefined;
}
