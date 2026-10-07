import { attribute, children } from './sheet';
import { fail } from './package-common';
import { visioFormulaCachedValue } from './formula';

export const cells = (shape: Element) =>
	new Map(children(shape, 'Cell').map((node) => [attribute(node, 'N') ?? '', node]));
export function numeric(node: Element | undefined, fallback?: number): number {
	if (!node && fallback !== undefined) return fallback;
	if (!node || node.hasAttribute('E'))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'A usable local numeric transform cache is required.');
	const value = visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value;
	if (Math.abs(value) > 1e6)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Geometry cache exceeds coordinate limits.');
	return value;
}
export function setCell(shape: Element, name: string, value: number, formula?: string): void {
	let node = cells(shape).get(name);
	if (!node) {
		node = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Cell');
		node.setAttribute('N', name);
		const before = Array.from(shape.childNodes).find(
			(child) => child.nodeType === 1 && (child as Element).localName !== 'Cell',
		);
		shape.insertBefore(node, before ?? null);
	}
	node.setAttribute('V', String(value));
	if (formula) node.setAttribute('F', formula);
	else if (attribute(node, 'F') !== 'No Formula') node.removeAttribute('F');
}
