/**
 * Small XML-object helpers the chart element format writers share: child
 * lookup by local name, ordered insertion against a schema sequence, and the
 * `a:srgbClr` + `a:alpha` colour choice the Format panes write.
 *
 * @module utils/chart-element-format-xml-util
 */
import type { XmlObject } from '../types';
import { chartColorHex } from './chart-color-choice';

export type LocalName = (key: string) => string;

export const EMU_PER_POINT = 12700;

/** Points to an EMU attribute value (never negative). */
export const pt = (value: number): string => String(Math.round(Math.max(0, value) * EMU_PER_POINT));

export function keyOf(node: XmlObject, name: string, localName: LocalName): string | undefined {
	return Object.keys(node).find((key) => !key.startsWith('@_') && localName(key) === name);
}

export function asObject(value: unknown): XmlObject | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

/** The child `name` of `node` as an object (an empty element reads as `{}`). */
export function childOf(
	node: XmlObject | undefined,
	name: string,
	localName: LocalName,
): XmlObject | undefined {
	if (!node) {
		return undefined;
	}
	const key = keyOf(node, name, localName);
	if (!key) {
		return undefined;
	}
	return asObject(node[key]) ?? {};
}

/** Remove the children named in `names`. */
export function removeChildren(
	node: XmlObject,
	names: readonly string[],
	localName: LocalName,
): void {
	for (const key of Object.keys(node)) {
		if (!key.startsWith('@_') && names.includes(localName(key))) {
			delete node[key];
		}
	}
}

/** Set `a:<name>` at its schema position in `order`, replacing an existing child of that name. */
export function setOrdered(
	node: XmlObject,
	name: string,
	value: XmlObject[string],
	order: readonly string[],
	localName: LocalName,
	prefix = 'a',
): void {
	const existing = keyOf(node, name, localName);
	if (existing) {
		node[existing] = value;
		return;
	}
	const rankOf = (key: string): number => {
		if (key.startsWith('@_')) {
			return -1;
		}
		const r = order.indexOf(localName(key));
		return r < 0 ? order.length : r;
	};
	// Stable sort: attributes first, then children in schema order (unknown ones last).
	const entries = [
		...Object.entries(node),
		[`${prefix}:${name}`, value] as [string, XmlObject[string]],
	]
		.map((entry, index) => ({ entry, index }))
		.sort((a, b) => rankOf(a.entry[0]) - rankOf(b.entry[0]) || a.index - b.index)
		.map(({ entry }) => entry);
	for (const key of Object.keys(node)) {
		delete node[key];
	}
	for (const [key, child] of entries) {
		node[key] = child;
	}
}

/** `a:alpha` for a transparency (0..100), or `undefined` when opaque. */
export function alphaOf(transparency: number | undefined): XmlObject | undefined {
	if (transparency === undefined || transparency <= 0) {
		return undefined;
	}
	const clamped = Math.min(100, Math.max(0, transparency));
	return { '@_val': String(Math.round((100 - clamped) * 1000)) };
}

/** `a:srgbClr` with an optional `a:alpha`. */
export function chartElementColorXml(color: string, transparency?: number): XmlObject {
	const clr: XmlObject = { '@_val': chartColorHex(color) };
	const alpha = alphaOf(transparency);
	if (alpha) {
		clr['a:alpha'] = alpha;
	}
	return { 'a:srgbClr': clr };
}

/** Update only the `a:alpha` of whatever colour choice `paint` wraps. */
export function setPaintAlpha(paint: XmlObject, transparency: number, localName: LocalName): void {
	const choiceKey = Object.keys(paint).find((key) => !key.startsWith('@_'));
	const choice = choiceKey ? asObject(paint[choiceKey]) : undefined;
	if (!choiceKey || !choice) {
		return;
	}
	const next: XmlObject = { ...choice };
	removeChildren(next, ['alpha'], localName);
	const alpha = alphaOf(transparency);
	if (alpha) {
		next['a:alpha'] = alpha;
	}
	paint[choiceKey] = next;
}

/** Replace every own key of `target` with `entries`, keeping their order. */
export function replaceEntries(target: XmlObject, entries: Array<[string, unknown]>): void {
	for (const key of Object.keys(target)) {
		delete target[key];
	}
	for (const [key, value] of entries) {
		target[key] = value as XmlObject[string];
	}
}
