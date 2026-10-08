// Shared helpers of the `c:chartSpace` writer: escaping, the canonical number form, `@val`
// children and the raw fragments (extension lists, kept sources) written under the part root.
import { stripDeclarations, type NamespaceBindings } from './xml-fragment';

/** What every writer function receives: the namespace bindings of the part root. */
export interface ChartWriteContext {
	bindings: NamespaceBindings;
}

export const escapeText = (value: string): string =>
	value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export const escapeAttribute = (value: string): string => escapeText(value).replace(/"/g, '&quot;');

/**
 * The canonical lexical form of a number the model holds as a value (`@val`, layout fractions):
 * the shortest round-trip decimal, never an exponent (`1e-7` becomes `0.0000001`), so Office reads
 * it as written. Values Office wrote with trailing zeros (`1.50`) come back as `1.5`.
 */
export function chartNumber(value: number): string {
	if (!Number.isFinite(value)) throw new Error(`Chart number ${value} is not finite`);
	const text = String(value);
	const match = /^(-?)(\d+)(?:\.(\d+))?e([+-]\d+)$/i.exec(text);
	if (!match) return text;
	const [, sign = '', whole = '', fraction = '', exponent = '0'] = match;
	const digits = whole + fraction;
	const point = whole.length + Number(exponent);
	const body =
		point <= 0
			? `0.${'0'.repeat(-point)}${digits}`
			: point >= digits.length
				? digits + '0'.repeat(point - digits.length)
				: `${digits.slice(0, point)}.${digits.slice(point)}`;
	return sign + body.replace(/^0+(?=\d)/, '');
}

/** `<c:local val="..."/>` for a defined value; booleans are written `1`/`0` as Office writes them. */
export function valXml(local: string, value: string | number | boolean | undefined): string {
	if (value === undefined) return '';
	const text =
		typeof value === 'boolean'
			? value
				? '1'
				: '0'
			: typeof value === 'number'
				? chartNumber(value)
				: escapeAttribute(value);
	return `<c:${local} val="${text}"/>`;
}

/** `<c:local>content</c:local>`, or the empty element when there is no content. */
export const elementXml = (local: string, content: string, attributes = ''): string =>
	content ? `<c:${local}${attributes}>${content}</c:${local}>` : `<c:${local}${attributes}/>`;

/** A raw fragment, with the declarations the part root already makes dropped. */
export const raw = (context: ChartWriteContext, xml: string | undefined): string =>
	xml ? stripDeclarations(xml, context.bindings) : '';

/** Structural equality of model values (key order ignored, absent and `undefined` alike). */
export function sameValue(left: unknown, right: unknown): boolean {
	if (left === right) return true;
	if (typeof left !== 'object' || typeof right !== 'object' || !left || !right) return false;
	if (Array.isArray(left) !== Array.isArray(right)) return false;
	const entries = (value: object) =>
		Object.entries(value).filter(([, item]) => item !== undefined) as [string, unknown][];
	const a = entries(left);
	const b = new Map(entries(right));
	return a.length === b.size && a.every(([key, item]) => b.has(key) && sameValue(item, b.get(key)));
}
