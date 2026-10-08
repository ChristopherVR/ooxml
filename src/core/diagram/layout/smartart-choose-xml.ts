/**
 * The one raw-XML representation the SmartArt layout interpreter walks: the
 * engine's document-order element tree (`engine/ordered-xml`). A layout
 * definition model keeps its raw slots (`rawXml`) in whatever form its reader
 * produced; a {@link RawXmlView} turns one slot into an element (pptx adapts
 * its `fast-xml-parser` objects at its own boundary).
 *
 * The walkers iterate children through {@link groupedChildren}: every child
 * with the same local name together, groups ordered by each name's first
 * appearance. That is the order the walkers were measured against (the pptx
 * object tree groups same-named siblings into one array per tag), so the
 * result is the same whether the element came from the object tree or from
 * the part text in document order.
 */

import type { OrderedXmlElement } from '../engine/ordered-xml';

export type { OrderedXmlElement } from '../engine/ordered-xml';

/** Turns a model's raw-XML slot into an element, or `undefined` when there is none. */
export type RawXmlView<R> = (raw: R | undefined) => OrderedXmlElement | undefined;

/** A view for models whose raw slots already are ordered elements. */
export const orderedXmlView: RawXmlView<OrderedXmlElement> = (raw) => raw;

const grouped = new WeakMap<OrderedXmlElement, readonly OrderedXmlElement[]>();

/** `element`'s children, same-named siblings grouped together in first-appearance order. */
export function groupedChildren(element: OrderedXmlElement): readonly OrderedXmlElement[] {
	const cached = grouped.get(element);
	if (cached) {
		return cached;
	}
	const groups = new Map<string, OrderedXmlElement[]>();
	for (const child of element.children) {
		const group = groups.get(child.name);
		if (group) {
			group.push(child);
		} else {
			groups.set(child.name, [child]);
		}
	}
	const result = [...groups.values()].flat();
	grouped.set(element, result);
	return result;
}

/** Children of `element` named `name` (local name), in document order. */
export function childrenNamed(
	element: OrderedXmlElement | undefined,
	name: string,
): OrderedXmlElement[] {
	return element ? element.children.filter((child) => child.name === name) : [];
}

/** The first child of `element` named `name` (local name). */
export function firstChildNamed(
	element: OrderedXmlElement | undefined,
	name: string,
): OrderedXmlElement | undefined {
	return element?.children.find((child) => child.name === name);
}

/** `dgm:param` children of a `dgm:alg` element as `{ type, value }`, skipping a missing `@type`. */
export function algorithmParameters(
	alg: OrderedXmlElement,
): { type: string; value: string | undefined }[] {
	const params: { type: string; value: string | undefined }[] = [];
	for (const param of childrenNamed(alg, 'param')) {
		const type = (param.attrs['type'] ?? '').trim();
		if (type) {
			params.push({ type, value: param.attrs['val'] });
		}
	}
	return params;
}
